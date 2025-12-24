// ===== AUTH.JS - Login & Auto-Save System =====
// Fitur: Login dengan username, auto-save form + foto, expiration 3 hari

const AUTH_CONFIG = {
    EXPIRATION_DAYS: 3,
    STORAGE_PREFIX: 'laporan_user_',
    SESSION_KEY: 'laporan_current_user',
    MAX_PHOTO_SIZE: 300 * 1024, // 300KB
    MAX_PHOTO_WIDTH: 1200
};

// ===== SESSION MANAGEMENT =====

function getCurrentUser() {
    return sessionStorage.getItem(AUTH_CONFIG.SESSION_KEY);
}

function setCurrentUser(username) {
    sessionStorage.setItem(AUTH_CONFIG.SESSION_KEY, username);
}

function clearCurrentUser() {
    sessionStorage.removeItem(AUTH_CONFIG.SESSION_KEY);
}

// ===== LOGIN / LOGOUT =====

function loginUser(username) {
    if (!username || username.trim() === '') {
        return { success: false, message: 'Username tidak boleh kosong' };
    }

    username = username.trim().toLowerCase();
    setCurrentUser(username);

    // Check for expired data first
    checkAndCleanExpiredData();

    // Try to load saved data
    const savedData = loadSavedData(username);

    if (savedData) {
        // Load data into form
        loadFormData(savedData.formData);
        return {
            success: true,
            message: `Selamat datang kembali, ${username}! Data terakhir berhasil dimuat.`,
            hasData: true,
            savedAt: savedData.savedAt
        };
    }

    return {
        success: true,
        message: `Selamat datang, ${username}! Anda bisa mulai mengisi form.`,
        hasData: false
    };
}

function logoutUser() {
    const username = getCurrentUser();
    if (username) {
        // Save current form data before logout
        saveCurrentFormData();
    }
    clearCurrentUser();
    return { success: true, message: 'Logout berhasil. Data Anda telah disimpan.' };
}

// ===== DATA STORAGE =====

function getStorageKey(username) {
    return AUTH_CONFIG.STORAGE_PREFIX + username;
}

function loadSavedData(username) {
    try {
        const key = getStorageKey(username);
        const data = localStorage.getItem(key);

        if (!data) return null;

        const parsed = JSON.parse(data);

        // Check if expired
        if (new Date(parsed.expiresAt) < new Date()) {
            localStorage.removeItem(key);
            return null;
        }

        return parsed;
    } catch (e) {
        console.error('Error loading saved data:', e);
        return null;
    }
}

function saveCurrentFormData() {
    const username = getCurrentUser();
    if (!username) return { success: false, message: 'Tidak ada user yang login' };

    try {
        const formData = collectFormData();
        const now = new Date();
        const expiresAt = new Date(now.getTime() + AUTH_CONFIG.EXPIRATION_DAYS * 24 * 60 * 60 * 1000);

        const dataToSave = {
            username: username,
            savedAt: now.toISOString(),
            expiresAt: expiresAt.toISOString(),
            formData: formData
        };

        const key = getStorageKey(username);
        localStorage.setItem(key, JSON.stringify(dataToSave));

        return {
            success: true,
            message: `Data berhasil disimpan! Akan expired pada ${expiresAt.toLocaleDateString('id-ID')}`
        };
    } catch (e) {
        console.error('Error saving form data:', e);
        if (e.name === 'QuotaExceededError') {
            return { success: false, message: 'Storage penuh! Hapus beberapa data lama.' };
        }
        return { success: false, message: 'Gagal menyimpan data: ' + e.message };
    }
}

// ===== COLLECT FORM DATA =====

function collectFormData() {
    const data = {
        bulan: document.getElementById('bulan')?.value || '',
        tahun: document.getElementById('tahun')?.value || '',
        jumlahLaporan: document.getElementById('jumlahLaporan')?.value || '1',
        laporans: {}
    };

    const jumlah = parseInt(data.jumlahLaporan);

    for (let i = 1; i <= jumlah; i++) {
        const laporanData = collectLaporanData(i);
        if (laporanData) {
            data.laporans[`laporan_${i}`] = laporanData;
        }
    }

    return data;
}

function collectLaporanData(index) {
    const prefix = `_${index}`;
    const data = {};

    // Basic fields
    const fields = [
        'tanggal', 'kategori', 'nama', 'usia', 'alamat',
        'beratBadan', 'tinggiBadan', 'tensi', 'keluhan', 'tindakLanjut', 'kader',
        'lila', 'lilaD', 'lika', 'lingkarPerut', 'lingkarPinggang'
    ];

    fields.forEach(field => {
        const element = document.getElementById(`${field}${prefix}`);
        if (element) {
            data[field] = element.value || '';
        }
    });

    // Checkboxes for balita
    const balitaCheckboxes = document.querySelectorAll(`input[name="pengukuranBalita${prefix}"]:checked`);
    data.pengukuranBalita = Array.from(balitaCheckboxes).map(cb => cb.value);

    // Checkboxes for dewasa
    const dewasaCheckboxes = document.querySelectorAll(`input[name="pengukuranDewasa${prefix}"]:checked`);
    data.pengukuranDewasa = Array.from(dewasaCheckboxes).map(cb => cb.value);

    // Photo - check if photo exists, actual data will be collected async in saveCurrentFormDataWithPhotos
    if (typeof compressedImageBlobs !== 'undefined' && compressedImageBlobs[index]) {
        data.hasPhoto = true;
        // Don't call async getPhotoBase64 here - it returns Promise
        // photoData will be set in saveCurrentFormDataWithPhotos
        data.photoData = null;
    } else {
        data.hasPhoto = false;
        data.photoData = null;
    }

    return data;
}

// ===== PHOTO HANDLING =====

async function getPhotoBase64(index) {
    if (typeof compressedImageBlobs === 'undefined' || !compressedImageBlobs[index]) {
        return null;
    }

    try {
        const blob = compressedImageBlobs[index];

        // Convert to WebP and resize if needed
        const webpBlob = await convertToWebP(blob, AUTH_CONFIG.MAX_PHOTO_WIDTH, AUTH_CONFIG.MAX_PHOTO_SIZE);

        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(webpBlob);
        });
    } catch (e) {
        console.error('Error converting photo:', e);
        return null;
    }
}

function convertToWebP(blob, maxWidth, maxSize) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');

            let { width, height } = img;
            if (width > maxWidth) {
                height = (height * maxWidth) / width;
                width = maxWidth;
            }

            canvas.width = width;
            canvas.height = height;
            ctx.drawImage(img, 0, 0, width, height);

            // Try different quality levels to get under maxSize
            let quality = 0.9;
            const tryConvert = () => {
                canvas.toBlob(
                    (resultBlob) => {
                        if (resultBlob && (resultBlob.size <= maxSize || quality <= 0.1)) {
                            resolve(resultBlob);
                        } else if (resultBlob) {
                            quality -= 0.1;
                            tryConvert();
                        } else {
                            reject(new Error('Failed to convert to WebP'));
                        }
                    },
                    'image/webp',
                    quality
                );
            };
            tryConvert();
        };
        img.onerror = () => reject(new Error('Failed to load image'));
        img.src = URL.createObjectURL(blob);
    });
}

// ===== SAVE WITH PHOTOS (ASYNC) =====

async function saveCurrentFormDataWithPhotos() {
    const username = getCurrentUser();
    if (!username) return { success: false, message: 'Tidak ada user yang login' };

    try {
        const formData = collectFormData();
        const jumlah = parseInt(formData.jumlahLaporan);

        // Process photos asynchronously
        for (let i = 1; i <= jumlah; i++) {
            if (formData.laporans[`laporan_${i}`]?.hasPhoto) {
                const photoBase64 = await getPhotoBase64(i);
                formData.laporans[`laporan_${i}`].photoData = photoBase64;
            }
        }

        const now = new Date();
        const expiresAt = new Date(now.getTime() + AUTH_CONFIG.EXPIRATION_DAYS * 24 * 60 * 60 * 1000);

        const dataToSave = {
            username: username,
            savedAt: now.toISOString(),
            expiresAt: expiresAt.toISOString(),
            formData: formData
        };

        const key = getStorageKey(username);
        localStorage.setItem(key, JSON.stringify(dataToSave));

        return {
            success: true,
            message: `Data + foto berhasil disimpan! Expired: ${expiresAt.toLocaleDateString('id-ID')}`
        };
    } catch (e) {
        console.error('Error saving form data with photos:', e);
        if (e.name === 'QuotaExceededError') {
            return { success: false, message: 'Storage penuh! Kurangi jumlah foto atau hapus data lama.' };
        }
        return { success: false, message: 'Gagal menyimpan: ' + e.message };
    }
}

// ===== LOAD FORM DATA =====

function loadFormData(formData) {
    if (!formData) return;

    // Basic fields
    if (formData.bulan) {
        const bulanEl = document.getElementById('bulan');
        if (bulanEl) bulanEl.value = formData.bulan;
    }

    if (formData.tahun) {
        const tahunEl = document.getElementById('tahun');
        if (tahunEl) tahunEl.value = formData.tahun;
    }

    if (formData.jumlahLaporan) {
        const jumlahEl = document.getElementById('jumlahLaporan');
        if (jumlahEl) {
            jumlahEl.value = formData.jumlahLaporan;
            // Trigger form generation
            if (typeof generateLaporanForms === 'function') {
                generateLaporanForms(parseInt(formData.jumlahLaporan));
            }
        }
    }

    // Wait for forms to generate, then load laporan data
    setTimeout(() => {
        if (formData.laporans) {
            Object.keys(formData.laporans).forEach((key, idx) => {
                const match = key.match(/laporan_(\d+)/);
                if (match) {
                    const index = parseInt(match[1]);
                    // Stagger loading to ensure DOM is ready
                    setTimeout(() => {
                        loadLaporanData(index, formData.laporans[key]);
                    }, idx * 100);
                }
            });
        }
    }, 300);
}

function loadLaporanData(index, data) {
    if (!data) return;

    const prefix = `_${index}`;

    // Basic fields
    const fields = [
        'tanggal', 'kategori', 'nama', 'usia', 'alamat',
        'beratBadan', 'tinggiBadan', 'tensi', 'keluhan', 'tindakLanjut', 'kader',
        'lila', 'lilaD', 'lika', 'lingkarPerut', 'lingkarPinggang'
    ];

    fields.forEach(field => {
        const element = document.getElementById(`${field}${prefix}`);
        if (element && data[field]) {
            element.value = data[field];
            // Trigger change event for kategori
            if (field === 'kategori') {
                element.dispatchEvent(new Event('change'));
            }
            // Trigger input event for usia
            if (field === 'usia') {
                element.dispatchEvent(new Event('input'));
            }
        }
    });

    // Load checkboxes
    if (data.pengukuranBalita && Array.isArray(data.pengukuranBalita)) {
        data.pengukuranBalita.forEach(value => {
            const checkbox = document.querySelector(`input[name="pengukuranBalita${prefix}"][value="${value}"]`);
            if (checkbox) {
                checkbox.checked = true;
                checkbox.dispatchEvent(new Event('change'));
            }
        });
    }

    if (data.pengukuranDewasa && Array.isArray(data.pengukuranDewasa)) {
        data.pengukuranDewasa.forEach(value => {
            const checkbox = document.querySelector(`input[name="pengukuranDewasa${prefix}"][value="${value}"]`);
            if (checkbox) {
                checkbox.checked = true;
                checkbox.dispatchEvent(new Event('change'));
            }
        });
    }

    // Load photo if exists - add delay to ensure DOM is ready after kategori change
    if (data.hasPhoto && data.photoData) {
        setTimeout(() => {
            loadPhotoFromBase64(index, data.photoData);
        }, 200);
    }
}

function loadPhotoFromBase64(index, base64Data) {
    if (!base64Data) {
        console.log(`No photo data for laporan ${index}`);
        return;
    }

    // Validate base64 data format
    if (typeof base64Data !== 'string' || !base64Data.startsWith('data:')) {
        console.error(`Invalid base64 data format for laporan ${index}:`, typeof base64Data);
        return;
    }

    console.log(`Loading photo for laporan ${index}, data length: ${base64Data.length}`);

    // Convert base64 to blob
    fetch(base64Data)
        .then(res => {
            if (!res.ok) {
                throw new Error(`Failed to fetch base64 data: ${res.status}`);
            }
            return res.blob();
        })
        .then(blob => {
            console.log(`Photo blob created for laporan ${index}, size: ${blob.size}`);

            // Store in compressedImageBlobs
            if (typeof compressedImageBlobs !== 'undefined') {
                compressedImageBlobs[index] = blob;

                // Also store array buffer for Word generation
                blob.arrayBuffer().then(buffer => {
                    if (typeof processedImageData !== 'undefined') {
                        processedImageData[index] = buffer;
                    }
                });

                // Update UI with retry logic for DOM elements
                const updateUI = () => {
                    const dropZone = document.getElementById(`dropZone_${index}`);
                    if (dropZone) {
                        dropZone.classList.add('has-file');
                        const content = dropZone.querySelector('.drop-zone-content');
                        if (content) {
                            content.innerHTML = `
                                <div class="upload-icon">✅</div>
                                <p><strong>Foto berhasil dimuat!</strong></p>
                                <p>Klik untuk mengganti foto</p>
                            `;
                        }

                        // Show preview
                        if (typeof showImagePreview === 'function') {
                            showImagePreview(index, 'saved_photo.webp');
                        }

                        // Update preview panel
                        if (typeof updatePreview === 'function') {
                            updatePreview();
                        }

                        console.log(`Photo UI updated for laporan ${index}`);
                        return true;
                    }
                    return false;
                };

                // Try to update UI immediately, retry if DOM not ready
                if (!updateUI()) {
                    console.log(`Waiting for DOM elements for laporan ${index}...`);
                    setTimeout(() => {
                        if (!updateUI()) {
                            setTimeout(updateUI, 500);
                        }
                    }, 200);
                }
            }
        })
        .catch(e => {
            console.error(`Error loading photo for laporan ${index}:`, e);
        });
}

// ===== CLEANUP EXPIRED DATA =====

function checkAndCleanExpiredData() {
    const now = new Date();
    const keysToRemove = [];

    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith(AUTH_CONFIG.STORAGE_PREFIX)) {
            try {
                const data = JSON.parse(localStorage.getItem(key));
                if (data.expiresAt && new Date(data.expiresAt) < now) {
                    keysToRemove.push(key);
                }
            } catch (e) {
                // Invalid data, remove it
                keysToRemove.push(key);
            }
        }
    }

    keysToRemove.forEach(key => {
        localStorage.removeItem(key);
        console.log(`Removed expired data: ${key}`);
    });

    return keysToRemove.length;
}

// ===== DELETE USER DATA =====

function deleteUserData(username) {
    if (!username) username = getCurrentUser();
    if (!username) return { success: false, message: 'Tidak ada user' };

    const key = getStorageKey(username);
    localStorage.removeItem(key);

    return { success: true, message: `Data untuk ${username} telah dihapus` };
}

// ===== UI HELPER =====

function showAuthMessage(message, isError = false) {
    const msgEl = document.getElementById('authMessage');
    if (msgEl) {
        msgEl.textContent = message;
        msgEl.className = 'auth-message ' + (isError ? 'error' : 'success');
        msgEl.style.display = 'block';
        setTimeout(() => {
            msgEl.style.display = 'none';
        }, 4000);
    }
}

function updateLoginUI() {
    const username = getCurrentUser();
    const loginPanel = document.getElementById('loginPanel');
    const userInfo = document.getElementById('userInfo');
    const usernameDisplay = document.getElementById('usernameDisplay');
    const formContainer = document.querySelector('.form-container');

    if (username) {
        // Logged in
        if (loginPanel) loginPanel.style.display = 'none';
        if (userInfo) userInfo.style.display = 'flex';
        if (usernameDisplay) usernameDisplay.textContent = username;
        if (formContainer) formContainer.classList.add('logged-in');
    } else {
        // Not logged in
        if (loginPanel) loginPanel.style.display = 'block';
        if (userInfo) userInfo.style.display = 'none';
        if (formContainer) formContainer.classList.remove('logged-in');
    }
}

// ===== INIT =====

function initAuth() {
    // Clean expired data on page load
    checkAndCleanExpiredData();

    // Check if already logged in (from session)
    const username = getCurrentUser();
    if (username) {
        const savedData = loadSavedData(username);
        if (savedData) {
            loadFormData(savedData.formData);
        }
    }

    updateLoginUI();
}

// Export for use in other scripts
window.AuthSystem = {
    login: loginUser,
    logout: logoutUser,
    save: saveCurrentFormDataWithPhotos,
    getCurrentUser: getCurrentUser,
    deleteData: deleteUserData,
    updateUI: updateLoginUI,
    init: initAuth,
    showMessage: showAuthMessage
};

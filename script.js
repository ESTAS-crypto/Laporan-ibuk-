let compressedImageBlobs = {}; // Objek untuk menyimpan blob gambar per laporan
let processedImageData = {}; // Objek untuk menyimpan data gambar per laporan

// Progress tracking
let currentProgress = 0;
const progressSteps = [
    { step: 10, text: "Memulai proses..." },
    { step: 20, text: "Memproses data form..." },
    { step: 40, text: "Mengompres foto..." },
    { step: 60, text: "Membuat dokumen Word..." },
    { step: 80, text: "Menambahkan gambar ke dokumen..." },
    { step: 95, text: "Menyiapkan download..." },
    { step: 100, text: "Selesai! Dokumen siap didownload." },
];

// WhatsApp Contact Configuration
const WHATSAPP_CONFIG = {
    phoneNumber: "62895385890629", // Ganti dengan nomor WA Anda (dengan kode negara tanpa +)
    defaultMessage: "Halo, saya butuh bantuan dengan aplikasi Generator Laporan Kunjungan Rumah.",
};

// Wait for docx library to load
function waitForDocx() {
    return new Promise((resolve, reject) => {
        let attempts = 0;
        const maxAttempts = 50; // 5 seconds max

        const checkDocx = () => {
            attempts++;
            if (typeof docx !== "undefined" && docx.Document) {
                resolve();
            } else if (attempts >= maxAttempts) {
                reject(new Error("Library docx gagal dimuat"));
            } else {
                setTimeout(checkDocx, 100);
            }
        };

        checkDocx();
    });
}

// WhatsApp Functions
function openWhatsApp(message = WHATSAPP_CONFIG.defaultMessage) {
    const encodedMessage = encodeURIComponent(message);
    const whatsappURL = `https://wa.me/${WHATSAPP_CONFIG.phoneNumber}?text=${encodedMessage}`;
    window.open(whatsappURL, "_blank");
}

// Initialize WhatsApp Float Button
function initWhatsAppFloat() {
    const whatsappFloat = document.getElementById("whatsappFloat");
    if (whatsappFloat) {
        whatsappFloat.addEventListener("click", function() {
            openWhatsApp();
        });
    }
}

// Image Modal Functions
function openImageModal(imageSrc) {
    const modal = document.getElementById("imageModal");
    const modalImage = document.getElementById("modalImage");

    modalImage.src = imageSrc;
    modal.style.display = "block";

    // Prevent body scroll when modal is open
    document.body.style.overflow = "hidden";
}

function closeImageModal() {
    const modal = document.getElementById("imageModal");
    modal.style.display = "none";

    // Restore body scroll
    document.body.style.overflow = "auto";
}

// Escape key to close modal
document.addEventListener("keydown", function(event) {
    if (event.key === "Escape") {
        closeImageModal();
    }
});

// Fungsi untuk update progress
function updateProgress(percentage, text) {
    const progressFill = document.getElementById("progressFill");
    const progressText = document.getElementById("progressText");
    const progressPercentage = document.getElementById("progressPercentage");

    if (progressFill && progressText && progressPercentage) {
        progressFill.style.width = percentage + "%";
        progressText.textContent = text;
        progressPercentage.textContent = percentage + "%";
    }
}

// Fungsi untuk simulasi progress bertahap
function simulateProgress(targetStep) {
    return new Promise((resolve) => {
        const currentStep = progressSteps.find((step) => step.step === targetStep);
        if (currentStep) {
            const interval = setInterval(() => {
                if (currentProgress < targetStep) {
                    currentProgress += 2;
                    updateProgress(currentProgress, currentStep.text);
                } else {
                    clearInterval(interval);
                    resolve();
                }
            }, 30);
        } else {
            resolve();
        }
    });
}

// Fungsi untuk mengompres gambar
function compressImage(file, maxWidth = 1200, maxSize = 300000) {
    return new Promise((resolve, reject) => {
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        const img = new Image();

        img.onload = function() {
            try {
                // Hitung ukuran baru
                let { width, height } = img;
                if (width > maxWidth) {
                    height = (height * maxWidth) / width;
                    width = maxWidth;
                }

                canvas.width = width;
                canvas.height = height;

                // Gambar ke canvas
                ctx.drawImage(img, 0, 0, width, height);

                // Kompres dengan kualitas bertahap
                let quality = 0.9;
                const tryCompress = () => {
                    canvas.toBlob(
                        (blob) => {
                            if (blob && (blob.size <= maxSize || quality <= 0.1)) {
                                resolve(blob);
                            } else if (blob) {
                                quality -= 0.1;
                                tryCompress();
                            } else {
                                reject(new Error("Failed to compress image"));
                            }
                        },
                        "image/jpeg",
                        quality
                    );
                };

                tryCompress();
            } catch (error) {
                reject(error);
            }
        };

        img.onerror = () => reject(new Error("Failed to load image"));
        img.src = URL.createObjectURL(file);
    });
}

// Fungsi untuk toggle field balita dan dewasa berdasarkan kategori
function toggleCategoryFields(laporanIndex, kategori) {
    const balitaFields = document.getElementById(`balitaFields_${laporanIndex}`);
    const dewasaFields = document.getElementById(`dewasaFields_${laporanIndex}`);
    const lilaInput = document.getElementById(`lila_${laporanIndex}`);
    const likaInput = document.getElementById(`lika_${laporanIndex}`);
    const lingkarPerutBalita = document.getElementById(`lingkarPerut_${laporanIndex}`);

    // Fields untuk dewasa
    const pilihanPengukuran = document.getElementById(`pilihanPengukuran_${laporanIndex}`);
    const lilaDInput = document.getElementById(`lilaD_${laporanIndex}`);
    const lingkarPinggang = document.getElementById(`lingkarPinggang_${laporanIndex}`);
    const lilaOption = document.getElementById(`lilaOption_${laporanIndex}`);
    const pinggangOption = document.getElementById(`pinggangOption_${laporanIndex}`);

    if (kategori === "balita") {
        // Show balita fields, hide dewasa fields
        balitaFields.style.display = "block";
        dewasaFields.style.display = "none";
        balitaFields.style.animation = "fadeIn 0.5s ease-in";

        // Set required for balita fields
        lilaInput.required = true;
        likaInput.required = true;
        lingkarPerutBalita.required = true;

        // Reset dewasa fields
        pilihanPengukuran.required = false;
        lilaDInput.required = false;
        lingkarPinggang.required = false;

        // Clear dewasa field values
        pilihanPengukuran.value = "";
        lilaDInput.value = "";
        lingkarPinggang.value = "";
        lilaOption.style.display = "none";
        pinggangOption.style.display = "none";
    } else if (kategori === "dewasa") {
        // Show dewasa fields, hide balita fields
        dewasaFields.style.display = "block";
        balitaFields.style.display = "none";
        dewasaFields.style.animation = "fadeIn 0.5s ease-in";

        // Set required for dewasa fields
        pilihanPengukuran.required = true;

        // Reset balita fields
        lilaInput.required = false;
        likaInput.required = false;
        lingkarPerutBalita.required = false;

        // Clear balita field values
        lilaInput.value = "";
        likaInput.value = "";
        lingkarPerutBalita.value = "";
    } else {
        // Hide both fields if no category selected
        balitaFields.style.display = "none";
        dewasaFields.style.display = "none";

        // Set all as not required
        lilaInput.required = false;
        likaInput.required = false;
        lingkarPerutBalita.required = false;
        pilihanPengukuran.required = false;
        lilaDInput.required = false;
        lingkarPinggang.required = false;

        // Clear all values
        lilaInput.value = "";
        likaInput.value = "";
        lingkarPerutBalita.value = "";
        pilihanPengukuran.value = "";
        lilaDInput.value = "";
        lingkarPinggang.value = "";
        lilaOption.style.display = "none";
        pinggangOption.style.display = "none";
    }
}

// Fungsi untuk toggle opsi pengukuran dewasa
function togglePengukuranOption(laporanIndex, pilihan) {
    const lilaOption = document.getElementById(`lilaOption_${laporanIndex}`);
    const pinggangOption = document.getElementById(`pinggangOption_${laporanIndex}`);
    const lilaDInput = document.getElementById(`lilaD_${laporanIndex}`);
    const lingkarPinggang = document.getElementById(`lingkarPinggang_${laporanIndex}`);

    // Reset semua
    lilaOption.style.display = "none";
    pinggangOption.style.display = "none";
    lilaDInput.required = false;
    lingkarPinggang.required = false;
    lilaDInput.value = "";
    lingkarPinggang.value = "";

    if (pilihan === "lila") {
        lilaOption.style.display = "block";
        lilaOption.style.animation = "fadeIn 0.3s ease-in";
        lilaDInput.required = true;
    } else if (pilihan === "pinggang") {
        pinggangOption.style.display = "block";
        pinggangOption.style.animation = "fadeIn 0.3s ease-in";
        lingkarPinggang.required = true;
    }
}

// Fungsi untuk setup drag and drop
function setupDragAndDrop(laporanIndex) {
    const dropZone = document.getElementById(`dropZone_${laporanIndex}`);
    const fileInput = document.getElementById(`foto_${laporanIndex}`);

    if (!dropZone || !fileInput) return;

    // Click handler untuk drop zone
    dropZone.addEventListener("click", () => {
        fileInput.click();
    });

    // Prevent default drag behaviors
    ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    // Highlight drop zone when item is dragged over it
    ["dragenter", "dragover"].forEach((eventName) => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add("dragover");
        });
    });

    ["dragleave", "drop"].forEach((eventName) => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove("dragover");
        });
    });

    // Handle dropped files
    dropZone.addEventListener("drop", (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileSelection(files[0], laporanIndex);
        }
    });

    // Handle file input change
    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFileSelection(e.target.files[0], laporanIndex);
        }
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
}

// Fungsi untuk handle file selection (drag/drop atau click)
async function handleFileSelection(file, laporanIndex) {
    const dropZone = document.getElementById(`dropZone_${laporanIndex}`);
    const previewDiv = document.getElementById(`photoPreview_${laporanIndex}`);

    // Validasi file
    if (!file.type.startsWith("image/")) {
        showFileError(dropZone, "File harus berupa gambar (JPG, PNG, JPEG)");
        return;
    }

    if (file.size > 10 * 1024 * 1024) {
        // 10MB
        showFileError(dropZone, "Ukuran file terlalu besar (maksimal 10MB)");
        return;
    }

    // Show loading state
    dropZone.classList.add("loading");

    try {
        // Kompres gambar
        compressedImageBlobs[laporanIndex] = await compressImage(file);

        // Convert to ArrayBuffer for docx
        processedImageData[laporanIndex] = await compressedImageBlobs[
            laporanIndex
        ].arrayBuffer();

        // Update drop zone state
        dropZone.classList.remove("loading", "error");
        dropZone.classList.add("has-file");

        // Update drop zone content
        const content = dropZone.querySelector(".drop-zone-content");
        content.innerHTML = `
            <div class="upload-icon">✅</div>
            <p><strong>Foto berhasil diupload!</strong></p>
            <p>Klik untuk mengganti foto</p>
        `;

        // Show preview
        showImagePreview(laporanIndex, file.name);
    } catch (error) {
        console.error("Error processing image:", error);
        showFileError(dropZone, "Error memproses gambar");

        // Clean up
        delete compressedImageBlobs[laporanIndex];
        delete processedImageData[laporanIndex];
    }
}

// Fungsi untuk menampilkan error pada file
function showFileError(dropZone, message) {
    dropZone.classList.remove("loading", "dragover", "has-file");
    dropZone.classList.add("error");

    const content = dropZone.querySelector(".drop-zone-content");
    content.innerHTML = `
        <div class="upload-icon">❌</div>
        <p><strong>Error: ${message}</strong></p>
        <p>Klik untuk coba lagi</p>
    `;

    // Remove error state after 3 seconds
    setTimeout(() => {
        dropZone.classList.remove("error");
        resetDropZoneContent(dropZone);
    }, 3000);
}

// Fungsi untuk reset konten drop zone
function resetDropZoneContent(dropZone) {
    const content = dropZone.querySelector(".drop-zone-content");
    content.innerHTML = `
        <div class="upload-icon">📷</div>
        <p><strong>Drag & Drop foto di sini</strong><br>atau klik untuk memilih file</p>
        <p class="file-types">Format: JPG, PNG, JPEG (Max 10MB)</p>
    `;
}

// Fungsi untuk menampilkan preview gambar dengan modal
function showImagePreview(laporanIndex, fileName) {
    const previewDiv = document.getElementById(`photoPreview_${laporanIndex}`);

    if (compressedImageBlobs[laporanIndex]) {
        const imageUrl = URL.createObjectURL(compressedImageBlobs[laporanIndex]);
        const img = document.createElement("img");
        img.src = imageUrl;

        // Add click event to open modal
        img.onclick = () => openImageModal(imageUrl);

        const container = document.createElement("div");
        container.className = "photo-container";

        const removeBtn = document.createElement("button");
        removeBtn.className = "remove-photo-btn";
        removeBtn.innerHTML = "×";
        removeBtn.title = "Hapus foto";
        removeBtn.type = "button";
        removeBtn.onclick = (e) => {
            e.stopPropagation();
            removePhoto(laporanIndex);
        };

        container.appendChild(img);
        container.appendChild(removeBtn);

        const info = document.createElement("div");
        info.className = "file-info";
        info.innerHTML = `📸 ${fileName} - ${(
            compressedImageBlobs[laporanIndex].size / 1024
        ).toFixed(1)} KB<br><small>Klik gambar untuk memperbesar</small>`;

        previewDiv.innerHTML = "";
        previewDiv.appendChild(container);
        previewDiv.appendChild(info);
    }
}

// Fungsi untuk menghapus foto
function removePhoto(laporanIndex) {
    const dropZone = document.getElementById(`dropZone_${laporanIndex}`);
    const previewDiv = document.getElementById(`photoPreview_${laporanIndex}`);
    const fileInput = document.getElementById(`foto_${laporanIndex}`);

    // Reset drop zone
    dropZone.classList.remove("has-file", "error", "loading");
    resetDropZoneContent(dropZone);

    // Clear preview
    previewDiv.innerHTML = "";

    // Reset file input
    fileInput.value = "";

    // Clean up data
    delete compressedImageBlobs[laporanIndex];
    delete processedImageData[laporanIndex];
}

// Fungsi untuk membuat form laporan dinamis
function generateLaporanForms(jumlah) {
    const container = document.getElementById("laporanContainer");
    const currentCount = container.children.length;

    if (jumlah > currentCount) {
        // Tambah form baru
        for (let i = currentCount + 1; i <= jumlah; i++) {
            const laporanSection = document.createElement("div");
            laporanSection.className = "laporan-section";
            laporanSection.setAttribute("data-laporan", i);

            laporanSection.innerHTML = `
                <h3 class="laporan-title">Laporan ${i}</h3>
                
                <div class="form-group">
                    <label for="tanggal_${i}">Tanggal:</label>
                    <input type="date" id="tanggal_${i}" name="tanggal_${i}" required />
                </div>

                <!-- Kategori Usia -->
                <div class="form-group">
                    <label for="kategori_${i}">Kategori Usia:</label>
                    <select id="kategori_${i}" name="kategori_${i}" required>
                        <option value="">Pilih Kategori</option>
                        <option value="balita">Balita (0-5 tahun)</option>
                        <option value="dewasa">Dewasa (>5 tahun)</option>
                    </select>
                </div>

                <div class="row">
                    <div class="col">
                        <div class="form-group">
                            <label for="nama_${i}">Nama:</label>
                            <input type="text" id="nama_${i}" name="nama_${i}" required />
                        </div>
                    </div>
                    <div class="col">
                        <div class="form-group">
                            <label for="usia_${i}">Usia (tahun):</label>
                            <input type="number" id="usia_${i}" name="usia_${i}" required />
                        </div>
                    </div>
                </div>

                <div class="form-group">
                    <label for="alamat_${i}">Alamat:</label>
                    <input type="text" id="alamat_${i}" name="alamat_${i}" required />
                </div>

                <div class="row">
                    <div class="col">
                        <div class="form-group">
                            <label for="beratBadan_${i}">Berat Badan (kg):</label>
                            <input type="number" step="0.1" id="beratBadan_${i}" name="beratBadan_${i}" required />
                        </div>
                    </div>
                    <div class="col">
                        <div class="form-group">
                            <label for="tinggiBadan_${i}">Tinggi Badan (cm):</label>
                            <input type="number" id="tinggiBadan_${i}" name="tinggiBadan_${i}" required />
                        </div>
                    </div>
                </div>

                <!-- Field khusus balita (hidden by default) -->
                <div class="balita-fields" id="balitaFields_${i}" style="display: none;">
                    <div class="row">
                        <div class="col">
                            <div class="form-group">
                                <label for="lila_${i}">LILA - Lingkar Lengan Atas (cm):</label>
                                <input type="number" step="0.1" id="lila_${i}" name="lila_${i}" />
                            </div>
                        </div>
                        <div class="col">
                            <div class="form-group">
                                <label for="lika_${i}">LIKA - Lingkar Kepala (cm):</label>
                                <input type="number" step="0.1" id="lika_${i}" name="lika_${i}" />
                            </div>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="lingkarPerut_${i}">Lingkar Perut (cm):</label>
                        <input type="number" id="lingkarPerut_${i}" name="lingkarPerut_${i}" />
                    </div>
                </div>

                <!-- Field khusus dewasa (hidden by default) -->
                <div class="dewasa-fields" id="dewasaFields_${i}" style="display: none;">
                    <!-- Pilihan pengukuran untuk dewasa -->
                    <div class="form-group">
                        <label for="pilihanPengukuran_${i}">Pilih Jenis Pengukuran:</label>
                        <select id="pilihanPengukuran_${i}" name="pilihanPengukuran_${i}">
                            <option value="">Pilih Pengukuran</option>
                            <option value="lila">LILA - Lingkar Lengan Atas</option>
                            <option value="pinggang">Lingkar Pinggang</option>
                        </select>
                    </div>
                    
                    <!-- Field LILA untuk dewasa -->
                    <div class="pengukuran-option" id="lilaOption_${i}" style="display: none;">
                        <div class="form-group">
                            <label for="lilaD_${i}">LILA - Lingkar Lengan Atas (cm):</label>
                            <input type="number" step="0.1" id="lilaD_${i}" name="lilaD_${i}" />
                        </div>
                    </div>
                    
                    <!-- Field Lingkar Pinggang untuk dewasa -->
                    <div class="pengukuran-option" id="pinggangOption_${i}" style="display: none;">
                        <div class="form-group">
                            <label for="lingkarPinggang_${i}">Lingkar Pinggang (cm):</label>
                            <input type="number" step="0.1" id="lingkarPinggang_${i}" name="lingkarPinggang_${i}" />
                        </div>
                    </div>
                </div>

                <div class="form-group">
                    <label for="tensi_${i}">Tensi:</label>
                    <input type="text" id="tensi_${i}" name="tensi_${i}" placeholder="contoh: 120/80" required />
                </div>

                <div class="form-group">
                    <label for="keluhan_${i}">Keluhan / Permasalahan:</label>
                    <textarea id="keluhan_${i}" name="keluhan_${i}" required></textarea>
                </div>

                <div class="form-group">
                    <label for="tindakLanjut_${i}">Tindak Lanjut:</label>
                    <textarea id="tindakLanjut_${i}" name="tindakLanjut_${i}" required></textarea>
                </div>

                <div class="form-group">
                    <label for="kader_${i}">Kader yang melakukan kunjungan rumah:</label>
                    <input type="text" id="kader_${i}" name="kader_${i}" required />
                </div>

                <div class="form-group">
                    <label for="foto_${i}">Foto Geotag (opsional):</label>
                    <div class="drop-zone" id="dropZone_${i}">
                        <div class="drop-zone-content">
                            <div class="upload-icon">📷</div>
                            <p><strong>Drag & Drop foto di sini</strong><br>atau klik untuk memilih file</p>
                            <p class="file-types">Format: JPG, PNG, JPEG (Max 10MB)</p>
                        </div>
                        <input type="file" id="foto_${i}" name="foto_${i}" accept="image/*" style="display: none;" />
                    </div>
                    <div class="file-info">
                        Foto akan dikompres otomatis (max 1200px, ≤300KB)
                    </div>
                    <div class="photo-preview" id="photoPreview_${i}"></div>
                </div>
            `;

            container.appendChild(laporanSection);

            // Set default tanggal
            const today = new Date().toISOString().split("T")[0];
            document.getElementById(`tanggal_${i}`).value = today;

            // Add event listeners
            addEventListeners(i);
        }
    } else if (jumlah < currentCount) {
        // Hapus form yang berlebih
        for (let i = currentCount; i > jumlah; i--) {
            const section = container.querySelector(`[data-laporan="${i}"]`);
            if (section) {
                container.removeChild(section);
            }
            // Hapus data gambar yang tersimpan
            delete compressedImageBlobs[i];
            delete processedImageData[i];
        }
    }
}

// Fungsi untuk menambahkan event listeners
function addEventListeners(laporanIndex) {
    // Event listener untuk kategori usia
    const kategoriSelect = document.getElementById(`kategori_${laporanIndex}`);
    kategoriSelect.addEventListener("change", function() {
        toggleCategoryFields(laporanIndex, this.value);
    });

    // Event listener untuk pilihan pengukuran dewasa
    const pilihanPengukuran = document.getElementById(`pilihanPengukuran_${laporanIndex}`);
    if (pilihanPengukuran) {
        pilihanPengukuran.addEventListener("change", function() {
            togglePengukuranOption(laporanIndex, this.value);
        });
    }

    // Setup drag and drop untuk foto
    setupDragAndDrop(laporanIndex);
}

// Event listener untuk perubahan jumlah laporan
document
    .getElementById("jumlahLaporan")
    .addEventListener("change", function(e) {
        const jumlah = parseInt(e.target.value);
        generateLaporanForms(jumlah);
    });

// Format tanggal sesuai contoh (dd-mm-yy)
function formatDate(dateString) {
    try {
        const date = new Date(dateString);
        const day = String(date.getDate()).padStart(2, "0");
        const month = String(date.getMonth() + 1).padStart(2, "0");
        const year = String(date.getFullYear()).slice(-2);
        return `${day}-${month}-${year}`;
    } catch (error) {
        console.error("Error formatting date:", error);
        return dateString;
    }
}

// Show error message
function showError(message) {
    const container = document.querySelector(".container");
    let errorDiv = document.getElementById("errorMessage");

    if (!errorDiv) {
        errorDiv = document.createElement("div");
        errorDiv.id = "errorMessage";
        errorDiv.className = "error";
        container.appendChild(errorDiv);
    }

    errorDiv.textContent = message;
    errorDiv.style.display = "block";

    setTimeout(() => {
        if (errorDiv) {
            errorDiv.style.display = "none";
        }
    }, 5000);
}

// Fungsi untuk mengambil data semua laporan
function getAllLaporanData() {
    const jumlahLaporan = parseInt(
        document.getElementById("jumlahLaporan").value
    );
    const allData = [];

    for (let i = 1; i <= jumlahLaporan; i++) {
        const kategori = document.getElementById(`kategori_${i}`).value;
        const data = {
            kategori: kategori,
            tanggal: document.getElementById(`tanggal_${i}`).value,
            nama: document.getElementById(`nama_${i}`).value,
            usia: document.getElementById(`usia_${i}`).value,
            alamat: document.getElementById(`alamat_${i}`).value,
            beratBadan: document.getElementById(`beratBadan_${i}`).value,
            tinggiBadan: document.getElementById(`tinggiBadan_${i}`).value,
            tensi: document.getElementById(`tensi_${i}`).value,
            keluhan: document.getElementById(`keluhan_${i}`).value,
            tindakLanjut: document.getElementById(`tindakLanjut_${i}`).value,
            kader: document.getElementById(`kader_${i}`).value,
            fotoData: processedImageData[i] || null,
        };

        // Tambahkan data berdasarkan kategori
        if (kategori === "balita") {
            data.lila = document.getElementById(`lila_${i}`).value;
            data.lika = document.getElementById(`lika_${i}`).value;
            data.lingkarPerut = document.getElementById(`lingkarPerut_${i}`).value;
        } else if (kategori === "dewasa") {
            const pilihanPengukuran = document.getElementById(`pilihanPengukuran_${i}`).value;
            if (pilihanPengukuran === "lila") {
                data.lilaD = document.getElementById(`lilaD_${i}`).value;
                data.pilihanPengukuran = "lila";
            } else if (pilihanPengukuran === "pinggang") {
                data.lingkarPinggang = document.getElementById(`lingkarPinggang_${i}`).value;
                data.pilihanPengukuran = "pinggang";
            }
        }

        allData.push(data);
    }

    return allData;
}

// Buat paragraf untuk data hasil dengan spacing yang sesuai
function createDataParagraphs(formData) {
    const paragraphs = [];

    // Data dasar dengan format yang sama persis seperti foto ke-1
    paragraphs.push(
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `Nama : ${formData.nama}`,
                    size: 22,
                }),
            ],
            spacing: { after: 100 },
        }),
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `Usia : ${formData.usia} th`,
                    size: 22,
                }),
            ],
            spacing: { after: 100 },
        }),
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `Alamat : ${formData.alamat}`,
                    size: 22,
                }),
            ],
            spacing: { after: 100 },
        }),
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `Bb : ${formData.beratBadan} Kg`,
                    size: 22,
                }),
            ],
            spacing: { after: 100 },
        }),
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `Tb : ${formData.tinggiBadan} cm`,
                    size: 22,
                }),
            ],
            spacing: { after: 100 },
        })
    );

    // Tambahkan data berdasarkan kategori
    if (formData.kategori === "balita" && formData.lila && formData.lika) {
        paragraphs.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `LILA : ${formData.lila} cm`,
                        size: 22,
                    }),
                ],
                spacing: { after: 100 },
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `LIKA : ${formData.lika} cm`,
                        size: 22,
                    }),
                ],
                spacing: { after: 100 },
            })
        );

        if (formData.lingkarPerut) {
            paragraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: `Lp : ${formData.lingkarPerut} cm`,
                            size: 22,
                        }),
                    ],
                    spacing: { after: 100 },
                })
            );
        }
    } else if (formData.kategori === "dewasa") {
        if (formData.pilihanPengukuran === "lila" && formData.lilaD) {
            paragraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: `LILA : ${formData.lilaD} cm`,
                            size: 22,
                        }),
                    ],
                    spacing: { after: 100 },
                })
            );
        } else if (formData.pilihanPengukuran === "pinggang" && formData.lingkarPinggang) {
            paragraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: `Lingkar Pinggang : ${formData.lingkarPinggang} cm`,
                            size: 22,
                        }),
                    ],
                    spacing: { after: 100 },
                })
            );
        }
    }

    paragraphs.push(
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `Tensi : ${formData.tensi}`,
                    size: 22,
                }),
            ],
            spacing: { after: 200 },
        })
    );

    // Spasi kosong sebelum keluhan
    paragraphs.push(
        new docx.Paragraph({
            children: [new docx.TextRun({ text: "", size: 22 })],
            spacing: { after: 100 },
        })
    );

    // Keluhan dengan format yang benar
    paragraphs.push(
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: "Keluhan / Permasalahan :",
                    size: 22,
                }),
            ],
            spacing: { after: 100 },
        })
    );

    // Split keluhan by line breaks dan buat paragraph untuk setiap baris
    const keluhanLines = formData.keluhan
        .split("\n")
        .filter((line) => line.trim() !== "");

    if (keluhanLines.length > 0) {
        paragraphs.push(
            new docx.Paragraph({
                children: [new docx.TextRun({ text: "", size: 22 })],
                spacing: { after: 80 },
            })
        );

        keluhanLines.forEach((line, index) => {
            paragraphs.push(
                new docx.Paragraph({
                    children: [new docx.TextRun({ text: line.trim(), size: 22 })],
                    spacing: { after: 80 },
                })
            );
        });
    } else {
        paragraphs.push(
            new docx.Paragraph({
                children: [new docx.TextRun({ text: "", size: 22 })],
                spacing: { after: 80 },
            })
        );
    }

    // Spasi kosong sebelum tindak lanjut
    paragraphs.push(
        new docx.Paragraph({
            children: [new docx.TextRun({ text: "", size: 22 })],
            spacing: { after: 120 },
        }),
        new docx.Paragraph({
            children: [new docx.TextRun({ text: "", size: 22 })],
            spacing: { after: 100 },
        })
    );

    // Tindak Lanjut dengan spacing yang sama
    paragraphs.push(
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: "Tindak Lanjut :",
                    size: 22,
                }),
            ],
            spacing: { after: 100 },
        })
    );

    // Split tindak lanjut by line breaks
    const tindakLanjutLines = formData.tindakLanjut
        .split("\n")
        .filter((line) => line.trim() !== "");

    if (tindakLanjutLines.length > 0) {
        paragraphs.push(
            new docx.Paragraph({
                children: [new docx.TextRun({ text: "", size: 22 })],
                spacing: { after: 80 },
            })
        );

        tindakLanjutLines.forEach((line, index) => {
            paragraphs.push(
                new docx.Paragraph({
                    children: [new docx.TextRun({ text: line.trim(), size: 22 })],
                    spacing: { after: 80 },
                })
            );
        });
    }

    return paragraphs;
}

// Buat cell foto dengan ukuran yang tepat
function createPhotoCell(fotoData) {
    if (!fotoData) {
        // Cell kosong dengan tinggi yang sesuai
        return [
            new docx.Paragraph({
                children: [new docx.TextRun({ text: "", size: 22 })],
                spacing: { after: 600 },
            }),
        ];
    }

    try {
        return [
            new docx.Paragraph({
                children: [
                    new docx.ImageRun({
                        data: fotoData,
                        transformation: {
                            width: 200,
                            height: 280,
                        },
                    }),
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: {
                    after: 120,
                    before: 40,
                },
            }),
        ];
    } catch (error) {
        console.error("Error adding image to document:", error);
        return [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "(Foto tidak dapat diproses)",
                        size: 22,
                    }),
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 600 },
            }),
        ];
    }
}

// Fungsi untuk membuat satu section laporan (satu halaman)
function createLaporanSection(laporan, bulan, tahun, laporanIndex) {
    const formattedDate = formatDate(laporan.tanggal);

    return {
        properties: {
            page: {
                margin: {
                    top: 720,
                    right: 720,
                    bottom: 720,
                    left: 720,
                },
            },
        },
        children: [
            // Header dengan judul laporan untuk setiap halaman
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "LAPORAN KUNJUNGAN RUMAH",
                        bold: true,
                        size: 26,
                    }),
                ],
                alignment: docx.AlignmentType.LEFT,
                spacing: { after: 80 },
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "PUSKESMAS BALAS KLUMPRIK",
                        bold: true,
                        size: 26,
                    }),
                ],
                alignment: docx.AlignmentType.LEFT,
                spacing: { after: 80 },
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `BULAN ${bulan} TAHUN ${tahun}`,
                        bold: true,
                        size: 26,
                    }),
                ],
                alignment: docx.AlignmentType.LEFT,
                spacing: { after: 200 },
            }),

            // Tabel untuk laporan individual
            new docx.Table({
                width: {
                    size: 100,
                    type: docx.WidthType.PERCENTAGE,
                },
                borders: {
                    top: {
                        style: docx.BorderStyle.SINGLE,
                        size: 3,
                        color: "000000",
                    },
                    bottom: {
                        style: docx.BorderStyle.SINGLE,
                        size: 3,
                        color: "000000",
                    },
                    left: {
                        style: docx.BorderStyle.SINGLE,
                        size: 3,
                        color: "000000",
                    },
                    right: {
                        style: docx.BorderStyle.SINGLE,
                        size: 3,
                        color: "000000",
                    },
                    insideHorizontal: {
                        style: docx.BorderStyle.SINGLE,
                        size: 3,
                        color: "000000",
                    },
                    insideVertical: {
                        style: docx.BorderStyle.SINGLE,
                        size: 3,
                        color: "000000",
                    },
                },
                rows: [
                    // Header row
                    new docx.TableRow({
                        children: [
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [
                                            new docx.TextRun({
                                                text: "NO",
                                                bold: true,
                                                size: 24,
                                            }),
                                        ],
                                        alignment: docx.AlignmentType.CENTER,
                                    }),
                                ],
                                width: { size: 8, type: docx.WidthType.PERCENTAGE },
                                verticalAlign: docx.VerticalAlign.CENTER,
                                shading: {
                                    fill: "d9d9d9",
                                },
                            }),
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [
                                            new docx.TextRun({
                                                text: "TANGGAL",
                                                bold: true,
                                                size: 24,
                                            }),
                                        ],
                                        alignment: docx.AlignmentType.CENTER,
                                    }),
                                ],
                                width: { size: 18, type: docx.WidthType.PERCENTAGE },
                                verticalAlign: docx.VerticalAlign.CENTER,
                                shading: {
                                    fill: "d9d9d9",
                                },
                            }),
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [
                                            new docx.TextRun({
                                                text: "HASIL",
                                                bold: true,
                                                size: 24,
                                            }),
                                        ],
                                        alignment: docx.AlignmentType.CENTER,
                                    }),
                                ],
                                width: { size: 50, type: docx.WidthType.PERCENTAGE },
                                verticalAlign: docx.VerticalAlign.CENTER,
                                shading: {
                                    fill: "d9d9d9",
                                },
                            }),
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [
                                            new docx.TextRun({
                                                text: "FOTO GEOTAG",
                                                bold: true,
                                                size: 24,
                                            }),
                                        ],
                                        alignment: docx.AlignmentType.CENTER,
                                    }),
                                ],
                                width: { size: 24, type: docx.WidthType.PERCENTAGE },
                                verticalAlign: docx.VerticalAlign.CENTER,
                                shading: {
                                    fill: "d9d9d9",
                                },
                            }),
                        ],
                    }),
                    // Data row untuk laporan ini
                    new docx.TableRow({
                        children: [
                            // NO cell - kosong untuk laporan individual
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [new docx.TextRun({ text: "", size: 22 })],
                                    }),
                                ],
                                verticalAlign: docx.VerticalAlign.TOP,
                            }),
                            // TANGGAL cell
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [
                                            new docx.TextRun({
                                                text: formattedDate,
                                                size: 22,
                                            }),
                                        ],
                                        alignment: docx.AlignmentType.CENTER,
                                    }),
                                ],
                                verticalAlign: docx.VerticalAlign.TOP,
                            }),
                            // HASIL cell dengan semua data
                            new docx.TableCell({
                                children: createDataParagraphs(laporan),
                                verticalAlign: docx.VerticalAlign.TOP,
                            }),
                            // FOTO GEOTAG cell
                            new docx.TableCell({
                                children: createPhotoCell(laporan.fotoData),
                                verticalAlign: docx.VerticalAlign.TOP,
                            }),
                        ],
                    }),
                    // Row untuk kader di bagian bawah
                    new docx.TableRow({
                        children: [
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [
                                            new docx.TextRun({
                                                text: "Kader yang melakukan kunjungan rumah : " +
                                                    laporan.kader,
                                                size: 22,
                                            }),
                                        ],
                                        alignment: docx.AlignmentType.LEFT,
                                        spacing: { after: 120 },
                                    }),
                                ],
                                columnSpan: 4,
                                verticalAlign: docx.VerticalAlign.CENTER,
                            }),
                        ],
                    }),
                ],
            }),
        ],
    };
}

// Generate laporan
document
    .getElementById("laporanForm")
    .addEventListener("submit", async function(e) {
        e.preventDefault();

        const generateBtn = document.getElementById("generateBtn");
        const progressContainer = document.getElementById("progressContainer");
        const loading = document.getElementById("loading");
        const success = document.getElementById("success");

        // Reset progress
        currentProgress = 0;

        generateBtn.disabled = true;
        progressContainer.style.display = "block";
        progressContainer.classList.add("fade-in");
        loading.style.display = "none";
        success.style.display = "none";

        try {
            // Wait for docx library to be ready
            await waitForDocx();

            // Step 1: Mulai proses
            await simulateProgress(10);

            // Step 2: Ambil data form
            await simulateProgress(20);
            const bulan = document.getElementById("bulan").value;
            const tahun = document.getElementById("tahun").value;
            const allLaporanData = getAllLaporanData();

            // Step 3: Proses foto jika ada
            const hasPhotos = Object.keys(processedImageData).length > 0;
            if (hasPhotos) {
                await simulateProgress(40);
            } else {
                currentProgress = 40;
                updateProgress(40, "Melewati kompresi foto...");
            }

            // Step 4: Buat dokumen Word
            await simulateProgress(60);

            // Step 5: Proses gambar untuk dokumen
            if (hasPhotos) {
                await simulateProgress(80);
            } else {
                currentProgress = 80;
                updateProgress(80, "Menyelesaikan dokumen...");
            }

            // Buat sections untuk setiap laporan (setiap laporan di halaman terpisah)
            const sections = allLaporanData.map((laporan, index) => {
                return createLaporanSection(laporan, bulan, tahun, index + 1);
            });

            // Buat dokumen Word dengan multiple sections (setiap section = halaman baru)
            const doc = new docx.Document({
                sections: sections,
            });

            // Step 6: Siapkan download
            await simulateProgress(95);

            // Generate dan download
            const blob = await docx.Packer.toBlob(doc);
            const fileName = `Laporan_Kunjungan_Rumah_${bulan}_${tahun}_${allLaporanData.length}laporan.docx`;

            // Step 7: Selesai
            await simulateProgress(100);

            // Download file
            saveAs(blob, fileName);

            // Show success message
            setTimeout(() => {
                progressContainer.style.display = "none";
                success.style.display = "block";
                success.classList.add("fade-in");

                setTimeout(() => {
                    success.style.display = "none";
                }, 3000);
            }, 500);
        } catch (error) {
            console.error("Error generating document:", error);
            progressContainer.style.display = "none";

            let errorMessage = "Terjadi kesalahan saat membuat dokumen: ";
            if (error.message.includes("docx")) {
                errorMessage +=
                    "Library docx gagal dimuat. Silakan refresh halaman dan coba lagi.";
            } else {
                errorMessage += error.message;
            }

            showError(errorMessage);
        } finally {
            generateBtn.disabled = false;
        }
    });

// Set tanggal hari ini sebagai default dan tahun saat ini
document.addEventListener("DOMContentLoaded", function() {
    const today = new Date();
    const formattedToday = today.toISOString().split("T")[0];

    // Set default untuk laporan pertama
    document.getElementById("tanggal_1").value = formattedToday;

    // Set tahun saat ini sebagai default
    const currentYear = today.getFullYear();
    document.getElementById("tahun").value = currentYear;

    // Set bulan saat ini sebagai default
    const currentMonth = today.getMonth();
    const bulanSelect = document.getElementById("bulan");
    const bulanOptions = [
        "",
        "JANUARI",
        "FEBRUARI",
        "MARET",
        "APRIL",
        "MEI",
        "JUNI",
        "JULI",
        "AGUSTUS",
        "SEPTEMBER",
        "OKTOBER",
        "NOVEMBER",
        "DESEMBER",
    ];
    bulanSelect.value = bulanOptions[currentMonth + 1];

    // Add event listeners untuk laporan pertama
    addEventListeners(1);

    // Initialize WhatsApp floating button
    initWhatsAppFloat();

    // Cek apakah library docx sudah dimuat
    setTimeout(() => {
        if (typeof docx === "undefined") {
            showError(
                "Peringatan: Library docx belum dimuat. Silakan refresh halaman."
            );
        }
    }, 2000);
});
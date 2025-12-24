// ===== SCRIPT.JS - IMPROVED DENGAN WORD LAYOUT YANG DIPERBAIKI =====

// ===== KONSTANTA & KONFIGURASI =====
const MAX_BALITA_TAHUN = 5;
const MAX_BALITA_BULAN = 60;

// ===== GLOBAL VARIABLES =====
let compressedImageBlobs = {};
let processedImageData = {};
let parsedAgeDataMap = {};

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

const WHATSAPP_CONFIG = {
    phoneNumber: "62895385890629",
    defaultMessage: "Halo, saya butuh bantuan dengan aplikasi Generator Laporan Kunjungan Rumah.",
};

// ===== SMART AGE PARSER (ENHANCED) =====
function parseAge(input) {
    if (!input || typeof input !== 'string') return null;

    // Normalize input: lowercase, replace comma with dot for decimals, trim spaces
    let normalized = input.trim().toLowerCase().replace(/\s+/g, ' ');
    const normalizedForParsing = normalized.replace(/,/g, '.');

    let totalBulan = 0;
    let totalTahun = 0;
    let detectedFormat = '';

    // Pattern 1: Kombinasi tahun dan bulan (e.g., "1 tahun 3 bulan", "1thn 3bln", "1th3bln")
    const comboPattern = /^(\d+\.?\d*)\s*(tahun|thn|th|year|yr)\s*(\d+\.?\d*)?\s*(bulan|bln|month|bl)?$/;
    const comboMatch = normalizedForParsing.match(comboPattern);

    if (comboMatch) {
        const tahunVal = parseFloat(comboMatch[1]) || 0;
        const bulanVal = parseFloat(comboMatch[3]) || 0;

        totalTahun = tahunVal + (bulanVal / 12);
        totalBulan = Math.round(tahunVal * 12 + bulanVal);
        detectedFormat = 'combo';
    }

    // Pattern 2: Hanya bulan (e.g., "10,4bulan", "3.5 bln", "24 month")
    if (!detectedFormat) {
        const bulanPattern = /^(\d+\.?\d*)\s*(bulan|bln|month|bl)$/;
        const bulanMatch = normalizedForParsing.match(bulanPattern);

        if (bulanMatch) {
            const nilai = parseFloat(bulanMatch[1]);
            if (!isNaN(nilai) && nilai >= 0) {
                totalBulan = Math.round(nilai);
                totalTahun = nilai / 12;
                detectedFormat = 'bulan';
            }
        }
    }

    // Pattern 3: Hanya tahun (e.g., "3,5thn", "2 tahun", "5 th", "1year")
    if (!detectedFormat) {
        const tahunPattern = /^(\d+\.?\d*)\s*(tahun|thn|th|year|yr)$/;
        const tahunMatch = normalizedForParsing.match(tahunPattern);

        if (tahunMatch) {
            const nilai = parseFloat(tahunMatch[1]);
            if (!isNaN(nilai) && nilai >= 0) {
                totalTahun = nilai;
                totalBulan = Math.round(nilai * 12);
                detectedFormat = 'tahun';
            }
        }
    }

    // Pattern 4: Angka saja tanpa satuan - akan ditentukan berdasarkan nilai
    // Jika <= 60, anggap bulan. Jika > 60, anggap tahun.
    if (!detectedFormat) {
        const numberOnlyPattern = /^(\d+\.?\d*)$/;
        const numberMatch = normalizedForParsing.match(numberOnlyPattern);

        if (numberMatch) {
            const nilai = parseFloat(numberMatch[1]);
            if (!isNaN(nilai) && nilai >= 0) {
                // Heuristic: jika <= 60, kemungkinan bulan; jika > 60, kemungkinan tahun
                if (nilai <= 60) {
                    totalBulan = Math.round(nilai);
                    totalTahun = nilai / 12;
                    detectedFormat = 'bulan_auto';
                } else {
                    // Nilai >60 anggap tahun (tidak mungkin usia >60 dalam bulan untuk balita)
                    totalTahun = nilai;
                    totalBulan = Math.round(nilai * 12);
                    detectedFormat = 'tahun_auto';
                }
            }
        }
    }

    // Pattern 5: Format dengan "b" atau "t" saja (e.g., "10b", "3t")
    if (!detectedFormat) {
        const shortPattern = /^(\d+\.?\d*)\s*([bt])$/;
        const shortMatch = normalizedForParsing.match(shortPattern);

        if (shortMatch) {
            const nilai = parseFloat(shortMatch[1]);
            const unit = shortMatch[2];
            if (!isNaN(nilai) && nilai >= 0) {
                if (unit === 'b') {
                    totalBulan = Math.round(nilai);
                    totalTahun = nilai / 12;
                    detectedFormat = 'bulan';
                } else {
                    totalTahun = nilai;
                    totalBulan = Math.round(nilai * 12);
                    detectedFormat = 'tahun';
                }
            }
        }
    }

    // Jika tidak ada format yang cocok, return null
    if (!detectedFormat) return null;

    // Validasi nilai tidak negatif
    if (totalBulan < 0 || totalTahun < 0) return null;

    // Format output string
    let formattedString;
    let displaySatuan;

    if (detectedFormat === 'combo') {
        const tahunBulat = Math.floor(totalTahun);
        const sisaBulan = totalBulan - (tahunBulat * 12);

        if (tahunBulat > 0 && sisaBulan > 0) {
            formattedString = `${tahunBulat} tahun ${sisaBulan} bulan`;
        } else if (tahunBulat > 0) {
            formattedString = `${tahunBulat} tahun`;
        } else {
            formattedString = `${sisaBulan} bulan`;
        }
        formattedString += ` (${totalBulan} bulan total)`;
        displaySatuan = 'combo';
    } else if (detectedFormat.startsWith('bulan')) {
        formattedString = `${totalBulan} bulan`;
        if (totalBulan >= 12) {
            const tahunFormatted = totalTahun.toFixed(1).replace('.', ',');
            formattedString += ` (${tahunFormatted} tahun)`;
        }
        displaySatuan = 'bulan';
        if (detectedFormat === 'bulan_auto') {
            formattedString += ' ⚠️ auto-detect';
        }
    } else { // tahun atau tahun_auto
        const nilaiFormatted = totalTahun.toString().replace('.', ',');
        formattedString = `${nilaiFormatted} tahun`;
        formattedString += ` (${totalBulan} bulan)`;
        displaySatuan = 'tahun';
        if (detectedFormat === 'tahun_auto') {
            formattedString += ' ⚠️ auto-detect';
        }
    }

    return {
        nilai: detectedFormat.startsWith('bulan') ? totalBulan : totalTahun,
        satuan: displaySatuan,
        formatted: formattedString,
        bulan: totalBulan,
        tahun: totalTahun,
        input: input,
        autoDetected: detectedFormat.endsWith('_auto')
    };
}

function validateBalitaAge(ageData) {
    if (!ageData) {
        return {
            valid: false,
            message: 'Format umur tidak valid'
        };
    }

    if (ageData.tahun > MAX_BALITA_TAHUN || ageData.bulan > MAX_BALITA_BULAN) {
        return {
            valid: false,
            message: `Umur balita tidak boleh melebihi ${MAX_BALITA_TAHUN} tahun (${MAX_BALITA_BULAN} bulan). Umur yang diinput: ${ageData.formatted}`
        };
    }

    return { valid: true, message: '' };
}

function handleUsiaInput(laporanIndex) {
    const usiaInput = document.getElementById(`usia_${laporanIndex}`);
    const kategoriSelect = document.getElementById(`kategori_${laporanIndex}`);
    const usiaParsedDiv = document.getElementById(`usiaParsed_${laporanIndex}`);
    const usiaErrorDiv = document.getElementById(`usiaError_${laporanIndex}`);

    if (!usiaInput || !usiaParsedDiv || !usiaErrorDiv) return;

    const input = usiaInput.value;
    const kategori = kategoriSelect ? kategoriSelect.value : '';

    usiaParsedDiv.classList.remove('show');
    usiaErrorDiv.classList.remove('show');
    delete parsedAgeDataMap[laporanIndex];

    if (!input) return;

    const ageData = parseAge(input);

    if (!ageData) {
        usiaErrorDiv.innerHTML = '❌ Format tidak valid.<br>Contoh: <b>10bulan</b>, <b>3,5thn</b>, <b>2tahun</b>, <b>1thn 3bulan</b>, atau <b>24</b> (angka saja)';
        usiaErrorDiv.classList.add('show');
        return;
    }

    if (kategori === 'balita') {
        const validation = validateBalitaAge(ageData);

        if (!validation.valid) {
            usiaErrorDiv.textContent = '❌ ' + validation.message;
            usiaErrorDiv.classList.add('show');
            return;
        }
    }

    parsedAgeDataMap[laporanIndex] = ageData;

    // Berikan warning jika menggunakan auto-detect
    if (ageData.autoDetected) {
        usiaParsedDiv.innerHTML = '⚠️ Umur diparsing otomatis sebagai: <b>' + ageData.formatted.replace(' ⚠️ auto-detect', '') + '</b><br><small style="color:#f39c12">Tambahkan satuan (bulan/thn) untuk hasil lebih akurat</small>';
    } else {
        usiaParsedDiv.innerHTML = '✓ Umur diparsing sebagai: <b>' + ageData.formatted + '</b>';
    }
    usiaParsedDiv.classList.add('show');

    updatePreview();
}

// ===== UTILITY FUNCTIONS =====

function waitForDocx() {
    return new Promise((resolve, reject) => {
        let attempts = 0;
        const maxAttempts = 50;
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

function openWhatsApp(message = WHATSAPP_CONFIG.defaultMessage) {
    const encodedMessage = encodeURIComponent(message);
    const whatsappURL = `https://wa.me/${WHATSAPP_CONFIG.phoneNumber}?text=${encodedMessage}`;
    window.open(whatsappURL, "_blank");
}

function initWhatsAppFloat() {
    const whatsappFloat = document.getElementById("whatsappFloat");
    if (whatsappFloat) {
        whatsappFloat.addEventListener("click", function () {
            openWhatsApp();
        });
    }
}

function openImageModal(imageSrc) {
    const modal = document.getElementById("imageModal");
    const modalImage = document.getElementById("modalImage");
    modalImage.src = imageSrc;
    modal.style.display = "block";
    document.body.style.overflow = "hidden";
}

function closeImageModal() {
    const modal = document.getElementById("imageModal");
    modal.style.display = "none";
    document.body.style.overflow = "auto";
}

document.addEventListener("keydown", function (event) {
    if (event.key === "Escape") {
        closeImageModal();
    }
});

function initPreviewPanel() {
    const toggleBtn = document.getElementById("togglePreview");
    const closeBtn = document.getElementById("closePreview");
    const previewPanel = document.getElementById("previewPanel");

    if (!toggleBtn || !closeBtn || !previewPanel) return;

    toggleBtn.addEventListener("click", () => {
        previewPanel.classList.toggle("active");
        if (previewPanel.classList.contains("active")) {
            updatePreview();
        }
    });

    closeBtn.addEventListener("click", () => {
        previewPanel.classList.remove("active");
    });

    document.addEventListener("click", (e) => {
        if (!previewPanel.contains(e.target) && e.target !== toggleBtn) {
            previewPanel.classList.remove("active");
        }
    });
}

function updatePreview() {
    const previewContent = document.getElementById("previewContent");
    const previewPanel = document.getElementById("previewPanel");

    if (!previewPanel || !previewPanel.classList.contains("active")) return;

    try {
        const bulan = document.getElementById("bulan").value;
        const tahun = document.getElementById("tahun").value;
        const jumlahLaporan = parseInt(
            document.getElementById("jumlahLaporan").value
        );

        if (!bulan || !tahun) {
            previewContent.innerHTML =
                '<p style="text-align: center; color: #999; padding: 20px;">Pilih bulan dan tahun terlebih dahulu</p>';
            return;
        }

        let allPreviewHTML = "";

        for (let i = 1; i <= jumlahLaporan; i++) {
            const tanggalElement = document.getElementById(`tanggal_${i}`);
            const namaElement = document.getElementById(`nama_${i}`);
            const usiaElement = document.getElementById(`usia_${i}`);
            const alamatElement = document.getElementById(`alamat_${i}`);
            const beratBadanElement = document.getElementById(`beratBadan_${i}`);
            const tinggiBadanElement = document.getElementById(`tinggiBadan_${i}`);
            const keluhanElement = document.getElementById(`keluhan_${i}`);
            const tindakLanjutElement = document.getElementById(`tindakLanjut_${i}`);
            const kaderElement = document.getElementById(`kader_${i}`);
            const tensiElement = document.getElementById(`tensi_${i}`);
            const kategoriElement = document.getElementById(`kategori_${i}`);

            if (!tanggalElement) continue;

            const tanggal = tanggalElement.value || "";
            const nama = namaElement ? namaElement.value : "";
            const usiaRaw = usiaElement ? usiaElement.value : "";
            const alamat = alamatElement ? alamatElement.value : "";
            const beratBadan = beratBadanElement ? beratBadanElement.value : "";
            const tinggiBadan = tinggiBadanElement ? tinggiBadanElement.value : "";
            const keluhan = keluhanElement ? keluhanElement.value : "";
            const tindakLanjut = tindakLanjutElement ? tindakLanjutElement.value : "";
            const kader = kaderElement ? kaderElement.value : "";
            const tensi = tensiElement ? tensiElement.value : "";
            const kategori = kategoriElement ? kategoriElement.value : "";

            let usiaDisplay = usiaRaw;
            if (parsedAgeDataMap[i]) {
                usiaDisplay = parsedAgeDataMap[i].formatted;
            }

            let measurements = "";
            const lilaElement = document.getElementById(`lila_${i}`);
            const lilaDElement = document.getElementById(`lilaD_${i}`);
            const likaElement = document.getElementById(`lika_${i}`);
            const lingkarPerutElement = document.getElementById(`lingkarPerut_${i}`);
            const lingkarPinggangElement = document.getElementById(`lingkarPinggang_${i}`);

            const lila = lilaElement ? lilaElement.value : "";
            const lilaD = lilaDElement ? lilaDElement.value : "";
            const lika = likaElement ? likaElement.value : "";
            const lingkarPerut = lingkarPerutElement ? lingkarPerutElement.value : "";
            const lingkarPinggang = lingkarPinggangElement ? lingkarPinggangElement.value : "";

            if (lila) measurements += `<p><strong>LILA:</strong> ${lila} cm</p>`;
            if (lilaD) measurements += `<p><strong>LILA:</strong> ${lilaD} cm</p>`;
            if (lika) measurements += `<p><strong>LIKA:</strong> ${lika} cm</p>`;
            if (lingkarPerut) measurements += `<p><strong>Lingkar Perut:</strong> ${lingkarPerut} cm</p>`;
            if (lingkarPinggang) measurements += `<p><strong>Lingkar Pinggang:</strong> ${lingkarPinggang} cm</p>`;

            const formattedDate = formatDate(tanggal);
            let fotoSrc = "";

            if (compressedImageBlobs[i]) {
                fotoSrc = URL.createObjectURL(compressedImageBlobs[i]);
            }

            if (i > 1) {
                allPreviewHTML += '<div class="preview-separator"></div>';
            }

            allPreviewHTML += `
        <div class="preview-page">
          <h1>LAPORAN KUNJUNGAN RUMAH</h1>
          <h1>PUSKESMAS BALAS KLUMPRIK</h1>
          <h2>BULAN ${bulan || "-"} TAHUN ${tahun || "-"}</h2>
          <table class="preview-table">
            <thead>
              <tr>
                <th class="col-no">NO</th>
                <th class="col-date">TANGGAL</th>
                <th class="col-result">HASIL</th>
                <th class="col-photo">FOTO GEOTAG</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="col-no">${i}</td>
                <td class="col-date">${formattedDate || "-"}</td>
                <td class="col-result">
                  <div class="result-info">
                    <p><strong>Nama:</strong> ${nama || "-"}</p>
                    <p><strong>Usia:</strong> ${usiaDisplay || "-"}</p>
                    <p><strong>Alamat:</strong> ${alamat || "-"}</p>
                    <p><strong>BB:</strong> ${beratBadan || "-"} Kg</p>
                    <p><strong>TB:</strong> ${tinggiBadan || "-"} cm</p>
                    ${measurements}
                    ${kategori !== 'balita' ? `<p><strong>Tensi:</strong> ${tensi || "-"}</p>` : ''}
                    <p style="margin-top: 8px;"><strong>Keluhan:</strong></p>
                    <p>${(keluhan || "-").replace(/\n/g, "<br>")}</p>
                    <p style="margin-top: 8px;"><strong>Tindak Lanjut:</strong></p>
                    <p>${(tindakLanjut || "-").replace(/\n/g, "<br>")}</p>
                  </div>
                </td>
                <td class="col-photo">
                  ${fotoSrc
                    ? `<img src="${fotoSrc}" class="preview-photo" alt="Foto" />`
                    : `<div class="no-photo">Belum ada foto</div>`
                }
                </td>
              </tr>
            </tbody>
          </table>
          <div class="preview-footer">
            <p><strong>Kader:</strong> ${kader || "-"}</p>
          </div>
        </div>
      `;
        }

        previewContent.innerHTML = allPreviewHTML;
    } catch (error) {
        console.error("Error updating preview:", error);
        previewContent.innerHTML =
            '<p style="text-align: center; color: #e74c3c; padding: 20px;">Terjadi kesalahan saat memuat preview</p>';
    }
}

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

function compressImage(file, maxWidth = 1200, maxSize = 300000) {
    return new Promise((resolve, reject) => {
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        const img = new Image();

        img.onload = function () {
            try {
                let { width, height } = img;
                if (width > maxWidth) {
                    height = (height * maxWidth) / width;
                    width = maxWidth;
                }

                canvas.width = width;
                canvas.height = height;
                ctx.drawImage(img, 0, 0, width, height);

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

function toggleCategoryFields(laporanIndex, kategori) {
    const balitaFields = document.getElementById(`balitaFields_${laporanIndex}`);
    const dewasaFields = document.getElementById(`dewasaFields_${laporanIndex}`);
    const tensiGroup = document.getElementById(`tensiGroup_${laporanIndex}`);
    const tensiInput = document.getElementById(`tensi_${laporanIndex}`);

    resetAllCheckboxes(laporanIndex);

    if (kategori === "balita") {
        if (balitaFields) balitaFields.style.display = "block";
        if (dewasaFields) dewasaFields.style.display = "none";
        if (tensiGroup) tensiGroup.style.display = "none";
        if (tensiInput) {
            tensiInput.removeAttribute('required');
            tensiInput.value = '';
        }
    } else if (kategori === "dewasa") {
        if (dewasaFields) dewasaFields.style.display = "block";
        if (balitaFields) balitaFields.style.display = "none";
        if (tensiGroup) tensiGroup.style.display = "block";
        if (tensiInput) tensiInput.setAttribute('required', 'required');
    } else {
        if (balitaFields) balitaFields.style.display = "none";
        if (dewasaFields) dewasaFields.style.display = "none";
        if (tensiGroup) tensiGroup.style.display = "block";
        if (tensiInput) tensiInput.setAttribute('required', 'required');
    }

    handleUsiaInput(laporanIndex);
    updatePreview();
}

function resetAllCheckboxes(laporanIndex) {
    const balitaCheckboxes = document.querySelectorAll(`input[name="pengukuranBalita_${laporanIndex}"]`);
    balitaCheckboxes.forEach((checkbox) => { checkbox.checked = false; });

    const dewasaCheckboxes = document.querySelectorAll(`input[name="pengukuranDewasa_${laporanIndex}"]`);
    dewasaCheckboxes.forEach((checkbox) => { checkbox.checked = false; });

    const inputFields = ["lila_", "lika_", "lingkarPerut_", "lilaD_", "lingkarPinggang_"];
    inputFields.forEach((fieldName) => {
        const field = document.getElementById(`${fieldName}${laporanIndex}`);
        if (field) field.value = "";
    });

    const allOptions = ["lilaBalitaOption_", "likaBalitaOption_", "lingkarPerutBalitaOption_", "lilaDewasaOption_", "lingkarPinggangDewasaOption_"];
    allOptions.forEach((optionName) => {
        const option = document.getElementById(`${optionName}${laporanIndex}`);
        if (option) option.style.display = "none";
    });
}

function toggleBalitaPengukuranOption(laporanIndex, checkboxValue, isChecked) {
    const optionMapping = {
        lila: "lilaBalitaOption_",
        lika: "likaBalitaOption_",
        lingkarPerut: "lingkarPerutBalitaOption_",
    };

    const optionId = optionMapping[checkboxValue] + laporanIndex;
    const option = document.getElementById(optionId);

    if (option) {
        if (isChecked) {
            option.style.display = "block";
        } else {
            option.style.display = "none";
            const input = option.querySelector("input");
            if (input) input.value = "";
        }
    }

    updatePreview();
}

function toggleDewasaPengukuranOption(laporanIndex, checkboxValue, isChecked) {
    const optionMapping = {
        lila: "lilaDewasaOption_",
        lingkarPinggang: "lingkarPinggangDewasaOption_",
    };

    const optionId = optionMapping[checkboxValue] + laporanIndex;
    const option = document.getElementById(optionId);

    if (option) {
        if (isChecked) {
            option.style.display = "block";
        } else {
            option.style.display = "none";
            const input = option.querySelector("input");
            if (input) input.value = "";
        }
    }

    updatePreview();
}

function setupDragAndDrop(laporanIndex) {
    const dropZone = document.getElementById(`dropZone_${laporanIndex}`);
    const fileInput = document.getElementById(`foto_${laporanIndex}`);

    if (!dropZone || !fileInput) return;

    dropZone.addEventListener("click", () => fileInput.click());

    ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    ["dragenter", "dragover"].forEach((eventName) => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add("dragover"));
    });

    ["dragleave", "drop"].forEach((eventName) => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove("dragover"));
    });

    dropZone.addEventListener("drop", (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) handleFileSelection(files[0], laporanIndex);
    });

    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) handleFileSelection(e.target.files[0], laporanIndex);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
}

async function handleFileSelection(file, laporanIndex) {
    const dropZone = document.getElementById(`dropZone_${laporanIndex}`);
    const previewDiv = document.getElementById(`photoPreview_${laporanIndex}`);

    if (!file.type.startsWith("image/")) {
        showFileError(dropZone, "File harus berupa gambar (JPG, PNG, JPEG)");
        return;
    }

    if (file.size > 10 * 1024 * 1024) {
        showFileError(dropZone, "Ukuran file terlalu besar (maksimal 10MB)");
        return;
    }

    dropZone.classList.add("loading");

    try {
        compressedImageBlobs[laporanIndex] = await compressImage(file);
        processedImageData[laporanIndex] = await compressedImageBlobs[laporanIndex].arrayBuffer();

        dropZone.classList.remove("loading", "error");
        dropZone.classList.add("has-file");

        const content = dropZone.querySelector(".drop-zone-content");
        content.innerHTML = `
      <div class="upload-icon">✅</div>
      <p><strong>Foto berhasil diupload!</strong></p>
      <p>Klik untuk mengganti foto</p>
    `;

        showImagePreview(laporanIndex, file.name);
        updatePreview();
    } catch (error) {
        console.error("Error processing image:", error);
        showFileError(dropZone, "Error memproses gambar");
        delete compressedImageBlobs[laporanIndex];
        delete processedImageData[laporanIndex];
    }
}

function showFileError(dropZone, message) {
    dropZone.classList.remove("loading", "dragover", "has-file");
    dropZone.classList.add("error");

    const content = dropZone.querySelector(".drop-zone-content");
    content.innerHTML = `
    <div class="upload-icon">❌</div>
    <p><strong>Error: ${message}</strong></p>
    <p>Klik untuk coba lagi</p>
  `;

    setTimeout(() => {
        dropZone.classList.remove("error");
        resetDropZoneContent(dropZone);
    }, 3000);
}

function resetDropZoneContent(dropZone) {
    const content = dropZone.querySelector(".drop-zone-content");
    content.innerHTML = `
    <div class="upload-icon">📷</div>
    <p><strong>Drag & Drop foto di sini</strong><br>atau klik untuk memilih file</p>
    <p class="file-types">Format: JPG, PNG, JPEG (Max 10MB)</p>
  `;
}

function showImagePreview(laporanIndex, fileName) {
    const previewDiv = document.getElementById(`photoPreview_${laporanIndex}`);

    if (compressedImageBlobs[laporanIndex]) {
        const imageUrl = URL.createObjectURL(compressedImageBlobs[laporanIndex]);
        const img = document.createElement("img");
        img.src = imageUrl;
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
        info.innerHTML = `📸 ${fileName} - ${(compressedImageBlobs[laporanIndex].size / 1024).toFixed(1)} KB<br><small>Klik gambar untuk memperbesar</small>`;

        previewDiv.innerHTML = "";
        previewDiv.appendChild(container);
        previewDiv.appendChild(info);
    }
}

function removePhoto(laporanIndex) {
    const dropZone = document.getElementById(`dropZone_${laporanIndex}`);
    const previewDiv = document.getElementById(`photoPreview_${laporanIndex}`);
    const fileInput = document.getElementById(`foto_${laporanIndex}`);

    dropZone.classList.remove("has-file", "error", "loading");
    resetDropZoneContent(dropZone);

    previewDiv.innerHTML = "";
    fileInput.value = "";

    delete compressedImageBlobs[laporanIndex];
    delete processedImageData[laporanIndex];

    updatePreview();
}

function generateLaporanForms(jumlah) {
    const container = document.getElementById("laporanContainer");
    const currentCount = container.children.length;

    if (jumlah > currentCount) {
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
              <label for="usia_${i}">Usia:</label>
              <input type="text" id="usia_${i}" name="usia_${i}" required 
                     placeholder="Contoh: 24, 10bulan, 3,5thn, 1thn 3bulan" />
              <div class="input-help">
                Format fleksibel: angka saja (auto-detect), angka+satuan (bulan/thn/tahun), atau kombinasi (1thn 3bulan)
              </div>
              <div id="usiaParsed_${i}" class="parsed-result"></div>
              <div id="usiaError_${i}" class="error-message"></div>
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

        <div class="balita-fields" id="balitaFields_${i}" style="display: none;">
          <div class="form-group">
            <label>Pilih Jenis Pengukuran (bisa pilih lebih dari 1):</label>
            <div class="checkbox-group">
              <div class="checkbox-item">
                <input type="checkbox" id="lilaBalitaCheck_${i}" name="pengukuranBalita_${i}" value="lila" />
                <label for="lilaBalitaCheck_${i}">LILA - Lingkar Lengan Atas</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="likaBalitaCheck_${i}" name="pengukuranBalita_${i}" value="lika" />
                <label for="likaBalitaCheck_${i}">LIKA - Lingkar Kepala</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="lingkarPerutBalitaCheck_${i}" name="pengukuranBalita_${i}" value="lingkarPerut" />
                <label for="lingkarPerutBalitaCheck_${i}">Lingkar Perut</label>
              </div>
            </div>
          </div>

          <div class="pengukuran-option" id="lilaBalitaOption_${i}" style="display: none">
            <div class="form-group">
              <label for="lila_${i}">LILA - Lingkar Lengan Atas (cm):</label>
              <input type="number" step="0.1" id="lila_${i}" name="lila_${i}" />
            </div>
          </div>
          <div class="pengukuran-option" id="likaBalitaOption_${i}" style="display: none">
            <div class="form-group">
              <label for="lika_${i}">LIKA - Lingkar Kepala (cm):</label>
              <input type="number" step="0.1" id="lika_${i}" name="lika_${i}" />
            </div>
          </div>
          <div class="pengukuran-option" id="lingkarPerutBalitaOption_${i}" style="display: none">
            <div class="form-group">
              <label for="lingkarPerut_${i}">Lingkar Perut (cm):</label>
              <input type="number" id="lingkarPerut_${i}" name="lingkarPerut_${i}" />
            </div>
          </div>
        </div>

        <div class="dewasa-fields" id="dewasaFields_${i}" style="display: none;">
          <div class="form-group">
            <label>Pilih Jenis Pengukuran (bisa pilih lebih dari 1):</label>
            <div class="checkbox-group">
              <div class="checkbox-item">
                <input type="checkbox" id="lilaDewasaCheck_${i}" name="pengukuranDewasa_${i}" value="lila" />
                <label for="lilaDewasaCheck_${i}">LILA - Lingkar Lengan Atas</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="lingkarPinggangDewasaCheck_${i}" name="pengukuranDewasa_${i}" value="lingkarPinggang" />
                <label for="lingkarPinggangDewasaCheck_${i}">Lingkar Pinggang</label>
              </div>
            </div>
          </div>

          <div class="pengukuran-option" id="lilaDewasaOption_${i}" style="display: none">
            <div class="form-group">
              <label for="lilaD_${i}">LILA - Lingkar Lengan Atas (cm):</label>
              <input type="number" step="0.1" id="lilaD_${i}" name="lilaD_${i}" />
            </div>
          </div>
          <div class="pengukuran-option" id="lingkarPinggangDewasaOption_${i}" style="display: none">
            <div class="form-group">
              <label for="lingkarPinggang_${i}">Lingkar Pinggang (cm):</label>
              <input type="number" step="0.1" id="lingkarPinggang_${i}" name="lingkarPinggang_${i}" />
            </div>
          </div>
        </div>

        <div class="form-group" id="tensiGroup_${i}">
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

            const today = new Date().toISOString().split("T")[0];
            document.getElementById(`tanggal_${i}`).value = today;

            addEventListeners(i);
        }
    } else if (jumlah < currentCount) {
        for (let i = currentCount; i > jumlah; i--) {
            const section = container.querySelector(`[data-laporan="${i}"]`);
            if (section) container.removeChild(section);
            delete compressedImageBlobs[i];
            delete processedImageData[i];
            delete parsedAgeDataMap[i];
        }
    }

    updatePreview();
}

function addEventListeners(laporanIndex) {
    const kategoriSelect = document.getElementById(`kategori_${laporanIndex}`);
    if (kategoriSelect) {
        kategoriSelect.addEventListener("change", function () {
            toggleCategoryFields(laporanIndex, this.value);
        });
    }

    const usiaInput = document.getElementById(`usia_${laporanIndex}`);
    if (usiaInput) {
        usiaInput.addEventListener("input", () => handleUsiaInput(laporanIndex));
        usiaInput.addEventListener("blur", () => handleUsiaInput(laporanIndex));
    }

    const balitaCheckboxes = document.querySelectorAll(`input[name="pengukuranBalita_${laporanIndex}"]`);
    balitaCheckboxes.forEach((checkbox) => {
        checkbox.addEventListener("change", function () {
            toggleBalitaPengukuranOption(laporanIndex, this.value, this.checked);
        });
    });

    const dewasaCheckboxes = document.querySelectorAll(`input[name="pengukuranDewasa_${laporanIndex}"]`);
    dewasaCheckboxes.forEach((checkbox) => {
        checkbox.addEventListener("change", function () {
            toggleDewasaPengukuranOption(laporanIndex, this.value, this.checked);
        });
    });

    setupDragAndDrop(laporanIndex);

    const inputs = [
        `tanggal_${laporanIndex}`,
        `nama_${laporanIndex}`,
        `alamat_${laporanIndex}`,
        `beratBadan_${laporanIndex}`,
        `tinggiBadan_${laporanIndex}`,
        `lila_${laporanIndex}`,
        `lilaD_${laporanIndex}`,
        `lika_${laporanIndex}`,
        `lingkarPerut_${laporanIndex}`,
        `lingkarPinggang_${laporanIndex}`,
        `tensi_${laporanIndex}`,
        `keluhan_${laporanIndex}`,
        `tindakLanjut_${laporanIndex}`,
        `kader_${laporanIndex}`,
    ];

    inputs.forEach((id) => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener("input", updatePreview);
            element.addEventListener("change", updatePreview);
        }
    });
}

document.getElementById("jumlahLaporan").addEventListener("change", function (e) {
    const jumlah = parseInt(e.target.value);
    generateLaporanForms(jumlah);
});

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
        if (errorDiv) errorDiv.style.display = "none";
    }, 5000);
}

function getAllLaporanData() {
    const jumlahLaporan = parseInt(document.getElementById("jumlahLaporan").value);
    const allData = [];

    for (let i = 1; i <= jumlahLaporan; i++) {
        const kategori = document.getElementById(`kategori_${i}`).value;
        const usiaRaw = document.getElementById(`usia_${i}`).value;

        if (!parsedAgeDataMap[i]) {
            showError(`Laporan ${i}: Format usia tidak valid atau belum diisi dengan benar`);
            return null;
        }

        const tensiInput = document.getElementById(`tensi_${i}`);
        let tensi = '';
        if (kategori === 'balita') {
            tensi = '-';
        } else {
            tensi = tensiInput ? tensiInput.value : '';
            if (!tensi) {
                showError(`Laporan ${i}: Tensi harus diisi untuk kategori ${kategori}`);
                return null;
            }
        }

        const data = {
            kategori: kategori,
            tanggal: document.getElementById(`tanggal_${i}`).value,
            nama: document.getElementById(`nama_${i}`).value,
            usia: parsedAgeDataMap[i].formatted,
            usiaRaw: usiaRaw,
            alamat: document.getElementById(`alamat_${i}`).value,
            beratBadan: document.getElementById(`beratBadan_${i}`).value,
            tinggiBadan: document.getElementById(`tinggiBadan_${i}`).value,
            tensi: tensi,
            keluhan: document.getElementById(`keluhan_${i}`).value,
            tindakLanjut: document.getElementById(`tindakLanjut_${i}`).value,
            kader: document.getElementById(`kader_${i}`).value,
            fotoData: processedImageData[i] || null,
            measurements: {},
        };

        if (kategori === "balita") {
            const balitaCheckboxes = document.querySelectorAll(`input[name="pengukuranBalita_${i}"]:checked`);
            balitaCheckboxes.forEach((checkbox) => {
                const value = checkbox.value;
                if (value === "lila") {
                    const lilaValue = document.getElementById(`lila_${i}`).value;
                    if (lilaValue) data.measurements.lila = lilaValue;
                } else if (value === "lika") {
                    const likaValue = document.getElementById(`lika_${i}`).value;
                    if (likaValue) data.measurements.lika = likaValue;
                } else if (value === "lingkarPerut") {
                    const lingkarPerutValue = document.getElementById(`lingkarPerut_${i}`).value;
                    if (lingkarPerutValue) data.measurements.lingkarPerut = lingkarPerutValue;
                }
            });
        } else if (kategori === "dewasa") {
            const dewasaCheckboxes = document.querySelectorAll(`input[name="pengukuranDewasa_${i}"]:checked`);
            dewasaCheckboxes.forEach((checkbox) => {
                const value = checkbox.value;
                if (value === "lila") {
                    const lilaValue = document.getElementById(`lilaD_${i}`).value;
                    if (lilaValue) data.measurements.lila = lilaValue;
                } else if (value === "lingkarPinggang") {
                    const lingkarPinggangValue = document.getElementById(`lingkarPinggang_${i}`).value;
                    if (lingkarPinggangValue) data.measurements.lingkarPinggang = lingkarPinggangValue;
                }
            });
        }

        allData.push(data);
    }

    return allData;
}

// ===== IMPROVED WORD DOCUMENT GENERATOR =====

/**
 * Membuat tabel dengan layout yang mirip preview
 * Menggunakan struktur 4 kolom: NO, TANGGAL, HASIL, FOTO GEOTAG
 */
function createImprovedWordTable(laporan, bulan, tahun, laporanIndex) {
    const formattedDate = formatDate(laporan.tanggal);

    // Build HASIL cell content dengan format yang rapi
    const hasilParagraphs = [];

    // Nama
    hasilParagraphs.push(new docx.Paragraph({
        children: [
            new docx.TextRun({ text: "Nama: ", bold: true, size: 20 }),
            new docx.TextRun({ text: laporan.nama, size: 20 })
        ],
        spacing: { after: 80, before: 0 },
    }));

    // Usia
    hasilParagraphs.push(new docx.Paragraph({
        children: [
            new docx.TextRun({ text: "Usia: ", bold: true, size: 20 }),
            new docx.TextRun({ text: laporan.usia, size: 20 })
        ],
        spacing: { after: 80, before: 0 },
    }));

    // Alamat
    hasilParagraphs.push(new docx.Paragraph({
        children: [
            new docx.TextRun({ text: "Alamat: ", bold: true, size: 20 }),
            new docx.TextRun({ text: laporan.alamat, size: 20 })
        ],
        spacing: { after: 80, before: 0 },
    }));

    // BB
    hasilParagraphs.push(new docx.Paragraph({
        children: [
            new docx.TextRun({ text: "BB: ", bold: true, size: 20 }),
            new docx.TextRun({ text: `${laporan.beratBadan} Kg`, size: 20 })
        ],
        spacing: { after: 80, before: 0 },
    }));

    // TB
    hasilParagraphs.push(new docx.Paragraph({
        children: [
            new docx.TextRun({ text: "TB: ", bold: true, size: 20 }),
            new docx.TextRun({ text: `${laporan.tinggiBadan} cm`, size: 20 })
        ],
        spacing: { after: 80, before: 0 },
    }));

    // Measurements
    if (laporan.measurements) {
        Object.keys(laporan.measurements).forEach((measurementType) => {
            const value = laporan.measurements[measurementType];
            let label = '';

            switch (measurementType) {
                case 'lila': label = 'LILA'; break;
                case 'lika': label = 'LIKA'; break;
                case 'lingkarPerut': label = 'Lingkar Perut'; break;
                case 'lingkarPinggang': label = 'Lingkar Pinggang'; break;
            }

            if (label && value) {
                hasilParagraphs.push(new docx.Paragraph({
                    children: [
                        new docx.TextRun({ text: `${label}: `, bold: true, size: 20 }),
                        new docx.TextRun({ text: `${value} cm`, size: 20 })
                    ],
                    spacing: { after: 80, before: 0 },
                }));
            }
        });
    }

    // Tensi (tidak tampil untuk balita)
    if (laporan.kategori !== 'balita' && laporan.tensi && laporan.tensi !== '-') {
        hasilParagraphs.push(new docx.Paragraph({
            children: [
                new docx.TextRun({ text: "Tensi: ", bold: true, size: 20 }),
                new docx.TextRun({ text: laporan.tensi, size: 20 })
            ],
            spacing: { after: 120, before: 0 },
        }));
    } else {
        hasilParagraphs.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: "", size: 20 })],
            spacing: { after: 120, before: 0 },
        }));
    }

    // Keluhan
    hasilParagraphs.push(new docx.Paragraph({
        children: [new docx.TextRun({ text: "Keluhan:", bold: true, size: 20 })],
        spacing: { after: 80, before: 0 },
    }));

    const keluhanLines = laporan.keluhan.split('\n').filter(line => line.trim() !== '');
    keluhanLines.forEach(line => {
        hasilParagraphs.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: line.trim(), size: 20 })],
            spacing: { after: 80, before: 0 },
        }));
    });

    // Spacing
    hasilParagraphs.push(new docx.Paragraph({
        children: [new docx.TextRun({ text: "", size: 20 })],
        spacing: { after: 120, before: 0 },
    }));

    // Tindak Lanjut
    hasilParagraphs.push(new docx.Paragraph({
        children: [new docx.TextRun({ text: "Tindak Lanjut:", bold: true, size: 20 })],
        spacing: { after: 80, before: 0 },
    }));

    const tindakLanjutLines = laporan.tindakLanjut.split('\n').filter(line => line.trim() !== '');
    tindakLanjutLines.forEach(line => {
        hasilParagraphs.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: line.trim(), size: 20 })],
            spacing: { after: 80, before: 0 },
        }));
    });

    // Buat foto cell dengan ukuran yang proporsional dan spacing atas
    const photoCellContent = laporan.fotoData
        ? [
            // Spacing paragraf kosong di atas foto untuk menurunkan posisi
            new docx.Paragraph({
                children: [new docx.TextRun({ text: '', size: 20 })],
                spacing: { after: 200 },
            }),
            new docx.Paragraph({
                children: [
                    new docx.ImageRun({
                        data: laporan.fotoData,
                        transformation: {
                            width: 120,  // Lebih kecil agar pas di kolom
                            height: 160  // Proporsional 3:4
                        },
                    }),
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 40, before: 0 },
            })
        ]
        : [new docx.Paragraph({
            children: [new docx.TextRun({ text: '', size: 20 })],
            spacing: { after: 400 },
        })];

    // Buat tabel dengan 2 baris: header dan data
    const table = new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 10, color: "000000" },
            bottom: { style: docx.BorderStyle.SINGLE, size: 10, color: "000000" },
            left: { style: docx.BorderStyle.SINGLE, size: 10, color: "000000" },
            right: { style: docx.BorderStyle.SINGLE, size: 10, color: "000000" },
            insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 10, color: "000000" },
            insideVertical: { style: docx.BorderStyle.SINGLE, size: 10, color: "000000" },
        },
        rows: [
            // Header row
            new docx.TableRow({
                height: { value: 600, rule: docx.HeightRule.ATLEAST },
                children: [
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [new docx.TextRun({ text: "NO", bold: true, size: 22 })],
                                alignment: docx.AlignmentType.CENTER,
                            }),
                        ],
                        width: { size: 6, type: docx.WidthType.PERCENTAGE },
                        verticalAlign: docx.VerticalAlign.CENTER,
                        shading: { fill: "E0E0E0" },
                        margins: {
                            top: 100,
                            bottom: 100,
                            left: 100,
                            right: 100,
                        },
                    }),
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [new docx.TextRun({ text: "TANGGAL", bold: true, size: 22 })],
                                alignment: docx.AlignmentType.CENTER,
                            }),
                        ],
                        width: { size: 14, type: docx.WidthType.PERCENTAGE },
                        verticalAlign: docx.VerticalAlign.CENTER,
                        shading: { fill: "E0E0E0" },
                        margins: {
                            top: 100,
                            bottom: 100,
                            left: 100,
                            right: 100,
                        },
                    }),
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [new docx.TextRun({ text: "HASIL", bold: true, size: 22 })],
                                alignment: docx.AlignmentType.CENTER,
                            }),
                        ],
                        width: { size: 55, type: docx.WidthType.PERCENTAGE },
                        verticalAlign: docx.VerticalAlign.CENTER,
                        shading: { fill: "E0E0E0" },
                        margins: {
                            top: 100,
                            bottom: 100,
                            left: 150,
                            right: 150,
                        },
                    }),
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [new docx.TextRun({ text: "FOTO GEOTAG", bold: true, size: 22 })],
                                alignment: docx.AlignmentType.CENTER,
                            }),
                        ],
                        width: { size: 25, type: docx.WidthType.PERCENTAGE },
                        verticalAlign: docx.VerticalAlign.CENTER,
                        shading: { fill: "E0E0E0" },
                        margins: {
                            top: 100,
                            bottom: 100,
                            left: 100,
                            right: 100,
                        },
                    }),
                ],
            }),
            // Data row
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [new docx.TextRun({ text: laporanIndex.toString(), size: 20 })],
                                alignment: docx.AlignmentType.CENTER,
                            }),
                        ],
                        verticalAlign: docx.VerticalAlign.TOP,
                        margins: {
                            top: 150,
                            bottom: 150,
                            left: 100,
                            right: 100,
                        },
                    }),
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [new docx.TextRun({ text: formattedDate, size: 20 })],
                                alignment: docx.AlignmentType.CENTER,
                            }),
                        ],
                        verticalAlign: docx.VerticalAlign.TOP,
                        margins: {
                            top: 150,
                            bottom: 150,
                            left: 100,
                            right: 100,
                        },
                    }),
                    new docx.TableCell({
                        children: hasilParagraphs,
                        verticalAlign: docx.VerticalAlign.TOP,
                        margins: {
                            top: 150,
                            bottom: 150,
                            left: 150,
                            right: 150,
                        },
                    }),
                    new docx.TableCell({
                        children: photoCellContent,
                        verticalAlign: docx.VerticalAlign.TOP,
                        margins: {
                            top: 150,
                            bottom: 150,
                            left: 100,
                            right: 100,
                        },
                    }),
                ],
            }),
        ],
    });

    return table;
}

/**
 * Membuat section untuk satu laporan dengan layout yang diperbaiki
 */
function createImprovedLaporanSection(laporan, bulan, tahun, laporanIndex) {
    return {
        properties: {
            page: {
                margin: { top: 720, right: 720, bottom: 720, left: 720 },
            },
        },
        children: [
            // Header
            new docx.Paragraph({
                children: [new docx.TextRun({ text: "LAPORAN KUNJUNGAN RUMAH", bold: true, size: 24 })],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 100 },
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: "PUSKESMAS BALAS KLUMPRIK", bold: true, size: 24 })],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 100 },
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: `BULAN ${bulan} TAHUN ${tahun}`, bold: true, size: 24 })],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 200 },
            }),

            // Tabel
            createImprovedWordTable(laporan, bulan, tahun, laporanIndex),

            // Footer dengan kader
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `Kader yang melakukan kunjungan rumah : ${laporan.kader}`,
                        size: 20,
                        bold: true,
                    }),
                ],
                alignment: docx.AlignmentType.LEFT,
                spacing: { before: 200, after: 120 },
                border: {
                    top: {
                        color: "000000",
                        space: 1,
                        style: docx.BorderStyle.SINGLE,
                        size: 6,
                    },
                },
            }),
        ],
    };
}

// ===== WORD FORM SUBMISSION =====

document.getElementById("laporanForm").addEventListener("submit", async function (e) {
    e.preventDefault();

    const generateBtn = document.getElementById("generateBtn");
    const progressContainer = document.getElementById("progressContainer");
    const loading = document.getElementById("loading");
    const success = document.getElementById("success");

    currentProgress = 0;

    generateBtn.disabled = true;
    progressContainer.style.display = "block";
    progressContainer.classList.add("fade-in");
    loading.style.display = "none";
    success.style.display = "none";

    try {
        await waitForDocx();
        await simulateProgress(10);
        await simulateProgress(20);

        const bulan = document.getElementById("bulan").value;
        const tahun = document.getElementById("tahun").value;
        const allLaporanData = getAllLaporanData();

        if (!allLaporanData) {
            throw new Error("Validasi data gagal");
        }

        const hasPhotos = Object.keys(processedImageData).length > 0;
        if (hasPhotos) {
            await simulateProgress(40);
        } else {
            currentProgress = 40;
            updateProgress(40, "Melewati kompresi foto...");
        }

        await simulateProgress(60);

        if (hasPhotos) {
            await simulateProgress(80);
        } else {
            currentProgress = 80;
            updateProgress(80, "Menyelesaikan dokumen...");
        }

        // Gunakan fungsi improved untuk membuat sections
        const sections = allLaporanData.map((laporan, index) => {
            return createImprovedLaporanSection(laporan, bulan, tahun, index + 1);
        });

        const doc = new docx.Document({ sections: sections });

        await simulateProgress(95);

        const blob = await docx.Packer.toBlob(doc);
        const fileName = `Laporan_Kunjungan_Rumah_${bulan}_${tahun}_${allLaporanData.length}laporan.docx`;

        await simulateProgress(100);

        saveAs(blob, fileName);

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
            errorMessage += "Library docx gagal dimuat. Silakan refresh halaman dan coba lagi.";
        } else {
            errorMessage += error.message;
        }

        showError(errorMessage);
    } finally {
        generateBtn.disabled = false;
    }
});

// ===== INITIALIZATION =====

document.addEventListener("DOMContentLoaded", function () {
    const today = new Date();
    const formattedToday = today.toISOString().split("T")[0];

    document.getElementById("tanggal_1").value = formattedToday;

    const currentYear = today.getFullYear();
    document.getElementById("tahun").value = currentYear;

    const currentMonth = today.getMonth();
    const bulanSelect = document.getElementById("bulan");
    const bulanOptions = [
        "", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
        "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER",
    ];
    bulanSelect.value = bulanOptions[currentMonth + 1];

    addEventListeners(1);
    initWhatsAppFloat();
    initPreviewPanel();

    bulanSelect.addEventListener("change", updatePreview);
    document.getElementById("tahun").addEventListener("input", updatePreview);

    setTimeout(() => {
        if (typeof docx === "undefined") {
            showError("Peringatan: Library docx belum dimuat. Silakan refresh halaman.");
        }
    }, 2000);

    console.log('✅ Generator Laporan dengan Improved Word Layout berhasil dimuat');
    console.log('📄 Word: Layout tabel 4 kolom identik dengan preview');
    console.log('📑 PDF: Menggunakan html2canvas + jsPDF');
    console.log('🔧 Smart Age Parser: 10,4bulan, 3,5thn, 2tahun');
});
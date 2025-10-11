let compressedImageBlobs = {};
let processedImageData = {};

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
  defaultMessage:
    "Halo, saya butuh bantuan dengan aplikasi Generator Laporan Kunjungan Rumah.",
};

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

    // Jika bulan atau tahun kosong
    if (!bulan || !tahun) {
      previewContent.innerHTML =
        '<p style="text-align: center; color: #999; padding: 20px;">Pilih bulan dan tahun terlebih dahulu</p>';
      return;
    }

    let allPreviewHTML = "";

    // Loop untuk setiap laporan
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

      // Skip jika element tidak ada
      if (!tanggalElement) continue;

      const tanggal = tanggalElement.value || "";
      const nama = namaElement ? namaElement.value : "";
      const usia = usiaElement ? usiaElement.value : "";
      const alamat = alamatElement ? alamatElement.value : "";
      const beratBadan = beratBadanElement ? beratBadanElement.value : "";
      const tinggiBadan = tinggiBadanElement ? tinggiBadanElement.value : "";
      const keluhan = keluhanElement ? keluhanElement.value : "";
      const tindakLanjut = tindakLanjutElement ? tindakLanjutElement.value : "";
      const kader = kaderElement ? kaderElement.value : "";
      const tensi = tensiElement ? tensiElement.value : "";

      let measurements = "";
      const lilaElement = document.getElementById(`lila_${i}`);
      const lilaDElement = document.getElementById(`lilaD_${i}`);
      const likaElement = document.getElementById(`lika_${i}`);
      const lingkarPerutElement = document.getElementById(`lingkarPerut_${i}`);
      const lingkarPinggangElement = document.getElementById(
        `lingkarPinggang_${i}`
      );

      const lila = lilaElement ? lilaElement.value : "";
      const lilaD = lilaDElement ? lilaDElement.value : "";
      const lika = likaElement ? likaElement.value : "";
      const lingkarPerut = lingkarPerutElement ? lingkarPerutElement.value : "";
      const lingkarPinggang = lingkarPinggangElement
        ? lingkarPinggangElement.value
        : "";

      if (lila) measurements += `<p><strong>LILA:</strong> ${lila} cm</p>`;
      if (lilaD) measurements += `<p><strong>LILA:</strong> ${lilaD} cm</p>`;
      if (lika) measurements += `<p><strong>LIKA:</strong> ${lika} cm</p>`;
      if (lingkarPerut)
        measurements += `<p><strong>Lingkar Perut:</strong> ${lingkarPerut} cm</p>`;
      if (lingkarPinggang)
        measurements += `<p><strong>Lingkar Pinggang:</strong> ${lingkarPinggang} cm</p>`;

      const formattedDate = formatDate(tanggal);
      let fotoSrc = "";

      if (compressedImageBlobs[i]) {
        fotoSrc = URL.createObjectURL(compressedImageBlobs[i]);
      }

      // Tambahkan separator jika bukan laporan pertama
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
                                <td class="col-date">${
                                  formattedDate || "-"
                                }</td>
                                <td class="col-result">
                                    <div class="result-info">
                                        <p><strong>Nama:</strong> ${
                                          nama || "-"
                                        }</p>
                                        <p><strong>Usia:</strong> ${
                                          usia || "-"
                                        } th</p>
                                        <p><strong>Alamat:</strong> ${
                                          alamat || "-"
                                        }</p>
                                        <p><strong>BB:</strong> ${
                                          beratBadan || "-"
                                        } Kg</p>
                                        <p><strong>TB:</strong> ${
                                          tinggiBadan || "-"
                                        } cm</p>
                                        ${measurements}
                                        <p><strong>Tensi:</strong> ${
                                          tensi || "-"
                                        }</p>
                                        <p style="margin-top: 8px;"><strong>Keluhan:</strong></p>
                                        <p>${(keluhan || "-").replace(
                                          /\n/g,
                                          "<br>"
                                        )}</p>
                                        <p style="margin-top: 8px;"><strong>Tindak Lanjut:</strong></p>
                                        <p>${(tindakLanjut || "-").replace(
                                          /\n/g,
                                          "<br>"
                                        )}</p>
                                    </div>
                                </td>
                                <td class="col-photo">
                                    ${
                                      fotoSrc
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

  resetAllCheckboxes(laporanIndex);

  if (kategori === "balita") {
    if (balitaFields) balitaFields.style.display = "block";
    if (dewasaFields) dewasaFields.style.display = "none";
  } else if (kategori === "dewasa") {
    if (dewasaFields) dewasaFields.style.display = "block";
    if (balitaFields) balitaFields.style.display = "none";
  } else {
    if (balitaFields) balitaFields.style.display = "none";
    if (dewasaFields) dewasaFields.style.display = "none";
  }

  updatePreview();
}

function resetAllCheckboxes(laporanIndex) {
  const balitaCheckboxes = document.querySelectorAll(
    `input[name="pengukuranBalita_${laporanIndex}"]`
  );
  balitaCheckboxes.forEach((checkbox) => {
    checkbox.checked = false;
  });

  const dewasaCheckboxes = document.querySelectorAll(
    `input[name="pengukuranDewasa_${laporanIndex}"]`
  );
  dewasaCheckboxes.forEach((checkbox) => {
    checkbox.checked = false;
  });

  const inputFields = [
    "lila_",
    "lika_",
    "lingkarPerut_",
    "lilaD_",
    "lingkarPinggang_",
  ];
  inputFields.forEach((fieldName) => {
    const field = document.getElementById(`${fieldName}${laporanIndex}`);
    if (field) {
      field.value = "";
    }
  });

  const allOptions = [
    "lilaBalitaOption_",
    "likaBalitaOption_",
    "lingkarPerutBalitaOption_",
    "lilaDewasaOption_",
    "lingkarPinggangDewasaOption_",
  ];
  allOptions.forEach((optionName) => {
    const option = document.getElementById(`${optionName}${laporanIndex}`);
    if (option) {
      option.style.display = "none";
    }
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
      if (input) {
        input.value = "";
      }
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
      if (input) {
        input.value = "";
      }
    }
  }

  updatePreview();
}

function setupDragAndDrop(laporanIndex) {
  const dropZone = document.getElementById(`dropZone_${laporanIndex}`);
  const fileInput = document.getElementById(`foto_${laporanIndex}`);

  if (!dropZone || !fileInput) return;

  dropZone.addEventListener("click", () => {
    fileInput.click();
  });

  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, preventDefaults, false);
    document.body.addEventListener(eventName, preventDefaults, false);
  });

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

  dropZone.addEventListener("drop", (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFileSelection(files[0], laporanIndex);
    }
  });

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
    processedImageData[laporanIndex] = await compressedImageBlobs[
      laporanIndex
    ].arrayBuffer();

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
    info.innerHTML = `📸 ${fileName} - ${(
      compressedImageBlobs[laporanIndex].size / 1024
    ).toFixed(1)} KB<br><small>Klik gambar untuk memperbesar</small>`;

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

      const today = new Date().toISOString().split("T")[0];
      document.getElementById(`tanggal_${i}`).value = today;

      addEventListeners(i);
    }
  } else if (jumlah < currentCount) {
    for (let i = currentCount; i > jumlah; i--) {
      const section = container.querySelector(`[data-laporan="${i}"]`);
      if (section) {
        container.removeChild(section);
      }
      delete compressedImageBlobs[i];
      delete processedImageData[i];
    }
  }

  // Update preview setelah generate form
  updatePreview();
}

function addEventListeners(laporanIndex) {
  const kategoriSelect = document.getElementById(`kategori_${laporanIndex}`);
  if (kategoriSelect) {
    kategoriSelect.addEventListener("change", function () {
      toggleCategoryFields(laporanIndex, this.value);
    });
  }

  const balitaCheckboxes = document.querySelectorAll(
    `input[name="pengukuranBalita_${laporanIndex}"]`
  );
  balitaCheckboxes.forEach((checkbox) => {
    checkbox.addEventListener("change", function () {
      toggleBalitaPengukuranOption(laporanIndex, this.value, this.checked);
    });
  });

  const dewasaCheckboxes = document.querySelectorAll(
    `input[name="pengukuranDewasa_${laporanIndex}"]`
  );
  dewasaCheckboxes.forEach((checkbox) => {
    checkbox.addEventListener("change", function () {
      toggleDewasaPengukuranOption(laporanIndex, this.value, this.checked);
    });
  });

  setupDragAndDrop(laporanIndex);

  const inputs = [
    `tanggal_${laporanIndex}`,
    `nama_${laporanIndex}`,
    `usia_${laporanIndex}`,
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

document
  .getElementById("jumlahLaporan")
  .addEventListener("change", function (e) {
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
    if (errorDiv) {
      errorDiv.style.display = "none";
    }
  }, 5000);
}

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
      measurements: {},
    };

    if (kategori === "balita") {
      const balitaCheckboxes = document.querySelectorAll(
        `input[name="pengukuranBalita_${i}"]:checked`
      );
      balitaCheckboxes.forEach((checkbox) => {
        const value = checkbox.value;
        if (value === "lila") {
          const lilaValue = document.getElementById(`lila_${i}`).value;
          if (lilaValue) data.measurements.lila = lilaValue;
        } else if (value === "lika") {
          const likaValue = document.getElementById(`lika_${i}`).value;
          if (likaValue) data.measurements.lika = likaValue;
        } else if (value === "lingkarPerut") {
          const lingkarPerutValue = document.getElementById(
            `lingkarPerut_${i}`
          ).value;
          if (lingkarPerutValue)
            data.measurements.lingkarPerut = lingkarPerutValue;
        }
      });
    } else if (kategori === "dewasa") {
      const dewasaCheckboxes = document.querySelectorAll(
        `input[name="pengukuranDewasa_${i}"]:checked`
      );
      dewasaCheckboxes.forEach((checkbox) => {
        const value = checkbox.value;
        if (value === "lila") {
          const lilaValue = document.getElementById(`lilaD_${i}`).value;
          if (lilaValue) data.measurements.lila = lilaValue;
        } else if (value === "lingkarPinggang") {
          const lingkarPinggangValue = document.getElementById(
            `lingkarPinggang_${i}`
          ).value;
          if (lingkarPinggangValue)
            data.measurements.lingkarPinggang = lingkarPinggangValue;
        }
      });
    }

    allData.push(data);
  }

  return allData;
}

function createDataParagraphs(formData) {
  const paragraphs = [];

  paragraphs.push(
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `Nama : ${formData.nama}`, size: 22 }),
      ],
      spacing: { after: 100 },
    }),
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `Usia : ${formData.usia} th`, size: 22 }),
      ],
      spacing: { after: 100 },
    }),
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `Alamat : ${formData.alamat}`, size: 22 }),
      ],
      spacing: { after: 100 },
    }),
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `Bb : ${formData.beratBadan} Kg`, size: 22 }),
      ],
      spacing: { after: 100 },
    }),
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `Tb : ${formData.tinggiBadan} cm`, size: 22 }),
      ],
      spacing: { after: 100 },
    })
  );

  if (formData.measurements) {
    Object.keys(formData.measurements).forEach((measurementType) => {
      const value = formData.measurements[measurementType];
      let label = "";

      switch (measurementType) {
        case "lila":
          label = "LILA";
          break;
        case "lika":
          label = "LIKA";
          break;
        case "lingkarPerut":
          label = "Lp";
          break;
        case "lingkarPinggang":
          label = "Lingkar Pinggang";
          break;
      }

      if (label && value) {
        paragraphs.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: `${label} : ${value} cm`, size: 22 }),
            ],
            spacing: { after: 100 },
          })
        );
      }
    });
  }

  paragraphs.push(
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `Tensi : ${formData.tensi}`, size: 22 }),
      ],
      spacing: { after: 200 },
    }),
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "", size: 22 })],
      spacing: { after: 100 },
    }),
    new docx.Paragraph({
      children: [
        new docx.TextRun({ text: "Keluhan / Permasalahan :", size: 22 }),
      ],
      spacing: { after: 100 },
    })
  );

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

    keluhanLines.forEach((line) => {
      paragraphs.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: line.trim(), size: 22 })],
          spacing: { after: 80 },
        })
      );
    });
  }

  paragraphs.push(
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "", size: 22 })],
      spacing: { after: 120 },
    }),
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "", size: 22 })],
      spacing: { after: 100 },
    }),
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "Tindak Lanjut :", size: 22 })],
      spacing: { after: 100 },
    })
  );

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

    tindakLanjutLines.forEach((line) => {
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

function createPhotoCell(fotoData) {
  if (!fotoData) {
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
        spacing: { after: 120, before: 40 },
      }),
    ];
  } catch (error) {
    console.error("Error adding image to document:", error);
    return [
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "(Foto tidak dapat diproses)", size: 22 }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 600 },
      }),
    ];
  }
}

function createLaporanSection(laporan, bulan, tahun, laporanIndex) {
  const formattedDate = formatDate(laporan.tanggal);

  return {
    properties: {
      page: {
        margin: { top: 720, right: 720, bottom: 720, left: 720 },
      },
    },
    children: [
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

      new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        borders: {
          top: { style: docx.BorderStyle.SINGLE, size: 3, color: "000000" },
          bottom: { style: docx.BorderStyle.SINGLE, size: 3, color: "000000" },
          left: { style: docx.BorderStyle.SINGLE, size: 3, color: "000000" },
          right: { style: docx.BorderStyle.SINGLE, size: 3, color: "000000" },
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
          new docx.TableRow({
            children: [
              new docx.TableCell({
                children: [
                  new docx.Paragraph({
                    children: [
                      new docx.TextRun({ text: "NO", bold: true, size: 24 }),
                    ],
                    alignment: docx.AlignmentType.CENTER,
                  }),
                ],
                width: { size: 8, type: docx.WidthType.PERCENTAGE },
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "d9d9d9" },
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
                shading: { fill: "d9d9d9" },
              }),
              new docx.TableCell({
                children: [
                  new docx.Paragraph({
                    children: [
                      new docx.TextRun({ text: "HASIL", bold: true, size: 24 }),
                    ],
                    alignment: docx.AlignmentType.CENTER,
                  }),
                ],
                width: { size: 50, type: docx.WidthType.PERCENTAGE },
                verticalAlign: docx.VerticalAlign.CENTER,
                shading: { fill: "d9d9d9" },
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
                shading: { fill: "d9d9d9" },
              }),
            ],
          }),
          new docx.TableRow({
            children: [
              new docx.TableCell({
                children: [
                  new docx.Paragraph({
                    children: [new docx.TextRun({ text: "", size: 22 })],
                  }),
                ],
                verticalAlign: docx.VerticalAlign.TOP,
              }),
              new docx.TableCell({
                children: [
                  new docx.Paragraph({
                    children: [
                      new docx.TextRun({ text: formattedDate, size: 22 }),
                    ],
                    alignment: docx.AlignmentType.CENTER,
                  }),
                ],
                verticalAlign: docx.VerticalAlign.TOP,
              }),
              new docx.TableCell({
                children: createDataParagraphs(laporan),
                verticalAlign: docx.VerticalAlign.TOP,
              }),
              new docx.TableCell({
                children: createPhotoCell(laporan.fotoData),
                verticalAlign: docx.VerticalAlign.TOP,
              }),
            ],
          }),
          new docx.TableRow({
            children: [
              new docx.TableCell({
                children: [
                  new docx.Paragraph({
                    children: [
                      new docx.TextRun({
                        text:
                          "Kader yang melakukan kunjungan rumah : " +
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

document
  .getElementById("laporanForm")
  .addEventListener("submit", async function (e) {
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

      const sections = allLaporanData.map((laporan, index) => {
        return createLaporanSection(laporan, bulan, tahun, index + 1);
      });

      const doc = new docx.Document({
        sections: sections,
      });

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

document.addEventListener("DOMContentLoaded", function () {
  const today = new Date();
  const formattedToday = today.toISOString().split("T")[0];

  document.getElementById("tanggal_1").value = formattedToday;

  const currentYear = today.getFullYear();
  document.getElementById("tahun").value = currentYear;

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

  addEventListeners(1);
  initWhatsAppFloat();
  initPreviewPanel();

  // Add event listeners untuk bulan dan tahun
  bulanSelect.addEventListener("change", updatePreview);
  document.getElementById("tahun").addEventListener("input", updatePreview);

  setTimeout(() => {
    if (typeof docx === "undefined") {
      showError(
        "Peringatan: Library docx belum dimuat. Silakan refresh halaman."
      );
    }
  }, 2000);
});

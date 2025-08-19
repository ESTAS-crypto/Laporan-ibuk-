let compressedImageBlob = null;

// Progress tracking
let currentProgress = 0;
const progressSteps = [
    { step: 10, text: "Memulai proses..." },
    { step: 20, text: "Memproses data form..." },
    { step: 40, text: "Mengompres foto..." },
    { step: 60, text: "Membuat dokumen Word..." },
    { step: 80, text: "Menambahkan gambar ke dokumen..." },
    { step: 95, text: "Menyiapkan download..." },
    { step: 100, text: "Selesai! Dokumen siap didownload." }
];

// Fungsi untuk update progress
function updateProgress(percentage, text) {
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const progressPercentage = document.getElementById('progressPercentage');

    if (progressFill && progressText && progressPercentage) {
        progressFill.style.width = percentage + '%';
        progressText.textContent = text;
        progressPercentage.textContent = percentage + '%';
    }
}

// Fungsi untuk simulasi progress bertahap
function simulateProgress(targetStep) {
    return new Promise((resolve) => {
        const currentStep = progressSteps.find(step => step.step === targetStep);
        if (currentStep) {
            const interval = setInterval(() => {
                if (currentProgress < targetStep) {
                    currentProgress += 1;
                    updateProgress(currentProgress, currentStep.text);
                } else {
                    clearInterval(interval);
                    resolve();
                }
            }, 50);
        } else {
            resolve();
        }
    });
}

// Fungsi untuk mengompres gambar
function compressImage(file, maxWidth = 1200, maxSize = 300000) {
    return new Promise((resolve) => {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const img = new Image();

        img.onload = function() {
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
                canvas.toBlob((blob) => {
                    if (blob.size <= maxSize || quality <= 0.1) {
                        resolve(blob);
                    } else {
                        quality -= 0.1;
                        tryCompress();
                    }
                }, 'image/jpeg', quality);
            };

            tryCompress();
        };

        img.src = URL.createObjectURL(file);
    });
}

// Preview foto
document.getElementById('foto').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    const previewDiv = document.getElementById('photoPreview');

    if (file) {
        const loading = document.getElementById('loading');
        loading.style.display = 'block';

        try {
            // Kompres gambar
            compressedImageBlob = await compressImage(file);

            // Preview
            const img = document.createElement('img');
            img.src = URL.createObjectURL(compressedImageBlob);
            previewDiv.innerHTML = '';
            previewDiv.appendChild(img);

            const info = document.createElement('div');
            info.className = 'file-info';
            info.textContent = `Ukuran setelah kompres: ${(compressedImageBlob.size / 1024).toFixed(1)} KB`;
            previewDiv.appendChild(info);
        } catch (error) {
            console.error('Error compressing image:', error);
            previewDiv.innerHTML = '<div style="color: red;">Error memproses gambar</div>';
        } finally {
            loading.style.display = 'none';
        }
    } else {
        previewDiv.innerHTML = '';
        compressedImageBlob = null;
    }
});

// Generate laporan
document.getElementById('laporanForm').addEventListener('submit', async function(e) {
    e.preventDefault();

    const generateBtn = document.getElementById('generateBtn');
    const progressContainer = document.getElementById('progressContainer');
    const loading = document.getElementById('loading');
    const success = document.getElementById('success');

    // Reset progress
    currentProgress = 0;

    generateBtn.disabled = true;
    progressContainer.style.display = 'block';
    progressContainer.classList.add('fade-in');
    loading.style.display = 'none';
    success.style.display = 'none';

    try {
        // Step 1: Mulai proses
        await simulateProgress(10);

        // Step 2: Ambil data form
        await simulateProgress(20);
        const formData = {
            tanggal: document.getElementById('tanggal').value,
            nama: document.getElementById('nama').value,
            usia: document.getElementById('usia').value,
            alamat: document.getElementById('alamat').value,
            beratBadan: document.getElementById('beratBadan').value,
            tinggiBadan: document.getElementById('tinggiBadan').value,
            lingkarPerut: document.getElementById('lingkarPerut').value,
            tensi: document.getElementById('tensi').value,
            keluhan: document.getElementById('keluhan').value,
            tindakLanjut: document.getElementById('tindakLanjut').value,
            kader: document.getElementById('kader').value
        };

        // Step 3: Proses foto jika ada
        if (compressedImageBlob) {
            await simulateProgress(40);
        } else {
            currentProgress = 40;
            updateProgress(40, "Melewati kompresi foto...");
        }

        // Step 4: Buat dokumen Word
        await simulateProgress(60);

        // Format tanggal sesuai contoh (dd-mm-yy)
        const dateObj = new Date(formData.tanggal);
        const formattedDate = dateObj.toLocaleDateString('id-ID', {
            day: '2-digit',
            month: '2-digit',
            year: '2-digit'
        });

        // Buat dokumen Word dengan format yang sama persis seperti contoh
        const doc = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 720, // 0.5 inch
                            right: 720, // 0.5 inch
                            bottom: 720, // 0.5 inch
                            left: 720, // 0.5 inch
                        },
                    },
                },
                children: [
                    // Header Table
                    new docx.Table({
                        width: {
                            size: 100,
                            type: docx.WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: { style: docx.BorderStyle.SINGLE, size: 1 },
                            bottom: { style: docx.BorderStyle.SINGLE, size: 1 },
                            left: { style: docx.BorderStyle.SINGLE, size: 1 },
                            right: { style: docx.BorderStyle.SINGLE, size: 1 },
                            insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1 },
                            insideVertical: { style: docx.BorderStyle.SINGLE, size: 1 }
                        },
                        rows: [
                            new docx.TableRow({
                                children: [
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: "LAPORAN\nKUNJUNGAN\nRUMAH",
                                                        bold: true,
                                                        size: 24
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.CENTER,
                                                spacing: { line: 360 }
                                            })
                                        ],
                                        width: { size: 20, type: docx.WidthType.PERCENTAGE },
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    }),
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: "PUSKESMAS\nBALAS\nKLUMPRIK\n\nBULAN\nAGUSTUS\nTAHUN 2025",
                                                        bold: true,
                                                        size: 24
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.CENTER,
                                                spacing: { line: 360 }
                                            })
                                        ],
                                        width: { size: 30, type: docx.WidthType.PERCENTAGE },
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    }),
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 24
                                                    })
                                                ],
                                                spacing: { line: 360 }
                                            })
                                        ],
                                        width: { size: 50, type: docx.WidthType.PERCENTAGE }
                                    })
                                ]
                            })
                        ]
                    }),

                    // Spacing
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: " " })],
                        spacing: { after: 240 }
                    }),

                    // Main Content Table
                    new docx.Table({
                        width: {
                            size: 100,
                            type: docx.WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: { style: docx.BorderStyle.SINGLE, size: 1 },
                            bottom: { style: docx.BorderStyle.SINGLE, size: 1 },
                            left: { style: docx.BorderStyle.SINGLE, size: 1 },
                            right: { style: docx.BorderStyle.SINGLE, size: 1 },
                            insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1 },
                            insideVertical: { style: docx.BorderStyle.SINGLE, size: 1 }
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
                                                        size: 22
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.CENTER
                                            })
                                        ],
                                        width: { size: 10, type: docx.WidthType.PERCENTAGE },
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    }),
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: "TANGGAL",
                                                        bold: true,
                                                        size: 22
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.CENTER
                                            })
                                        ],
                                        width: { size: 15, type: docx.WidthType.PERCENTAGE },
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    }),
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: "HASIL",
                                                        bold: true,
                                                        size: 22
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.CENTER
                                            })
                                        ],
                                        width: { size: 45, type: docx.WidthType.PERCENTAGE },
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    }),
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: "FOTO GEOTAG",
                                                        bold: true,
                                                        size: 22
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.CENTER
                                            })
                                        ],
                                        width: { size: 30, type: docx.WidthType.PERCENTAGE },
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    })
                                ]
                            }),
                            // Data row
                            new docx.TableRow({
                                children: [
                                    // NO cell (empty dengan space besar)
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 400 // Large space
                                                    })
                                                ]
                                            })
                                        ],
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    }),
                                    // TANGGAL cell
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: formattedDate,
                                                        size: 22
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.CENTER
                                            })
                                        ],
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    }),
                                    // HASIL cell
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Nama : ${formData.nama}`,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Usia : ${formData.usia} th`,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Alamat : ${formData.alamat}`,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Bb : ${formData.beratBadan} Kg`,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Tb : ${formData.tinggiBadan} cm`,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Lp : ${formData.lingkarPerut} cm`,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Tensi : ${formData.tensi}`,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 240 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: "Keluhan / Permasalahan :",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: formData.keluhan,
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 240 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: "Tindak Lanjut :",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 22
                                                    })
                                                ],
                                                spacing: { after: 120 }
                                            }),
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: formData.tindakLanjut,
                                                        size: 22
                                                    })
                                                ]
                                            })
                                        ],
                                        verticalAlign: docx.VerticalAlign.TOP
                                    }),
                                    // FOTO GEOTAG cell
                                    new docx.TableCell({
                                        children: await createPhotoCell(),
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    })
                                ]
                            }),
                            // Kader row
                            new docx.TableRow({
                                children: [
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: " ",
                                                        size: 22
                                                    })
                                                ]
                                            })
                                        ],
                                        columnSpan: 3
                                    }),
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({
                                                children: [
                                                    new docx.TextRun({
                                                        text: `Kader yang\nmelakukan\nkunjungan\nrumah :\n${formData.kader}`,
                                                        size: 22
                                                    })
                                                ],
                                                alignment: docx.AlignmentType.LEFT
                                            })
                                        ],
                                        verticalAlign: docx.VerticalAlign.CENTER
                                    })
                                ]
                            })
                        ]
                    })
                ]
            }]
        });

        // Step 5: Proses gambar untuk dokumen
        if (compressedImageBlob) {
            await simulateProgress(80);
        } else {
            currentProgress = 80;
            updateProgress(80, "Menyelesaikan dokumen...");
        }

        // Step 6: Siapkan download
        await simulateProgress(95);

        // Generate dan download
        const blob = await docx.Packer.toBlob(doc);
        const fileName = `Laporan_${formData.nama.replace(/\s+/g, '_')}_${formattedDate.replace(/\//g, '-')}.docx`;

        // Step 7: Selesai
        await simulateProgress(100);

        // Download file
        saveAs(blob, fileName);

        // Show success message
        setTimeout(() => {
            progressContainer.style.display = 'none';
            success.style.display = 'block';
            success.classList.add('fade-in');

            setTimeout(() => {
                success.style.display = 'none';
            }, 3000);
        }, 500);

    } catch (error) {
        console.error('Error generating document:', error);
        progressContainer.style.display = 'none';
        alert('Terjadi kesalahan saat membuat dokumen. Silakan coba lagi.');
    } finally {
        generateBtn.disabled = false;
    }
});

// Fungsi untuk membuat cell foto
async function createPhotoCell() {
    if (!compressedImageBlob) {
        return [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: " ",
                        size: 200
                    })
                ],
                spacing: { after: 240 }
            })
        ];
    }

    try {
        // Convert blob to array buffer
        const arrayBuffer = await compressedImageBlob.arrayBuffer();

        return [
            new docx.Paragraph({
                children: [
                    new docx.ImageRun({
                        data: arrayBuffer,
                        transformation: {
                            width: 167, // 2.31875 inches * 72 points/inch ≈ 167 points
                            height: 248, // 3.45 inches * 72 points/inch ≈ 248 points
                        }
                    })
                ],
                alignment: docx.AlignmentType.CENTER
            })
        ];
    } catch (error) {
        console.error('Error processing image for document:', error);
        return [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "(Foto tidak dapat diproses)",
                        size: 22
                    })
                ],
                alignment: docx.AlignmentType.CENTER
            })
        ];
    }
}

// Set tanggal hari ini sebagai default
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('tanggal').valueAsDate = new Date();
});
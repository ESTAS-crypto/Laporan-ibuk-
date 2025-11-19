// ===== STYLEPDF.JS - PDF GENERATOR =====
// File ini menangani pembuatan PDF dari preview HTML dengan style yang sama

/**
 * Fungsi utama untuk generate PDF
 * Menggunakan html2canvas untuk capture preview dan jsPDF untuk membuat PDF
 */
async function generatePDF() {
    const generateBtn = document.getElementById('generatePdfBtn');
    const progressContainer = document.getElementById('progressContainer');
    const loading = document.getElementById('loading');
    const success = document.getElementById('success');

    // Validasi form
    const form = document.getElementById('laporanForm');
    if (!form.checkValidity()) {
        form.reportValidity();
        return;
    }

    // Validasi data umur
    const allLaporanData = getAllLaporanData();
    if (!allLaporanData) {
        return;
    }

    try {
        // Disable button dan show progress
        generateBtn.disabled = true;
        progressContainer.style.display = 'block';
        loading.style.display = 'block';
        success.style.display = 'none';
        updateProgress(10, 'Mempersiapkan dokumen PDF...');

        // Get data
        const bulan = document.getElementById('bulan').value;
        const tahun = document.getElementById('tahun').value;
        const jumlahLaporan = parseInt(document.getElementById('jumlahLaporan').value);

        await simulateProgress(20);
        updateProgress(30, 'Membuat template PDF...');

        // Buat container sementara untuk render PDF
        const pdfContainer = document.createElement('div');
        pdfContainer.id = 'pdf-container';
        pdfContainer.style.position = 'absolute';
        pdfContainer.style.left = '-9999px';
        pdfContainer.style.top = '0';
        pdfContainer.style.width = '210mm'; // A4 width
        pdfContainer.style.background = 'white';
        document.body.appendChild(pdfContainer);

        await simulateProgress(40);

        // Generate HTML untuk setiap laporan
        for (let i = 1; i <= jumlahLaporan; i++) {
            const laporanData = allLaporanData[i - 1];
            const pageHtml = createPDFPage(laporanData, bulan, tahun, i);
            pdfContainer.appendChild(pageHtml);
        }

        await simulateProgress(60);
        updateProgress(70, 'Mengkonversi ke PDF...');

        // Tunggu sedikit agar semua gambar ter-render
        await new Promise(resolve => setTimeout(resolve, 500));

        // Buat PDF menggunakan jsPDF
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        await simulateProgress(80);

        // Capture setiap halaman
        const pages = pdfContainer.querySelectorAll('.pdf-page');
        for (let i = 0; i < pages.length; i++) {
            if (i > 0) {
                pdf.addPage();
            }

            // Capture halaman dengan html2canvas
            const canvas = await html2canvas(pages[i], {
                scale: 2,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff',
                width: pages[i].scrollWidth,
                height: pages[i].scrollHeight
            });

            const imgData = canvas.toDataURL('image/jpeg', 0.95);
            const imgWidth = 210; // A4 width in mm
            const imgHeight = (canvas.height * imgWidth) / canvas.width;

            pdf.addImage(imgData, 'JPEG', 0, 0, imgWidth, imgHeight);

            updateProgress(80 + (i + 1) * (15 / pages.length), `Memproses halaman ${i + 1}/${pages.length}...`);
        }

        await simulateProgress(95);

        // Hapus container sementara
        document.body.removeChild(pdfContainer);

        // Save PDF
        const fileName = `Laporan_Kunjungan_Rumah_${bulan}_${tahun}_${jumlahLaporan}laporan.pdf`;
        pdf.save(fileName);

        await simulateProgress(100);

        // Show success message
        setTimeout(() => {
            progressContainer.style.display = 'none';
            loading.style.display = 'none';
            success.style.display = 'block';
            success.classList.add('fade-in');

            setTimeout(() => {
                success.style.display = 'none';
            }, 3000);
        }, 500);

    } catch (error) {
        console.error('Error generating PDF:', error);
        progressContainer.style.display = 'none';
        loading.style.display = 'none';
        showError('Terjadi kesalahan saat membuat PDF: ' + error.message);
    } finally {
        generateBtn.disabled = false;
    }
}

/**
 * Membuat elemen HTML untuk satu halaman PDF
 * Style disesuaikan agar tampil identik dengan preview
 */
function createPDFPage(laporanData, bulan, tahun, laporanIndex) {
    const page = document.createElement('div');
    page.className = 'pdf-page';

    // Apply inline styles untuk memastikan tampilan konsisten
    page.style.cssText = `
        width: 210mm;
        min-height: 297mm;
        background: white;
        padding: 20px;
        box-sizing: border-box;
        font-family: Arial, sans-serif;
        page-break-after: always;
        position: relative;
    `;

    const formattedDate = formatDate(laporanData.tanggal);

    // Build measurements HTML
    let measurementsHtml = '';
    if (laporanData.measurements) {
        Object.keys(laporanData.measurements).forEach((measurementType) => {
            const value = laporanData.measurements[measurementType];
            let label = '';

            switch (measurementType) {
                case 'lila':
                    label = 'LILA';
                    break;
                case 'lika':
                    label = 'LIKA';
                    break;
                case 'lingkarPerut':
                    label = 'Lingkar Perut';
                    break;
                case 'lingkarPinggang':
                    label = 'Lingkar Pinggang';
                    break;
            }

            if (label && value) {
                measurementsHtml += `<p style="margin: 4px 0;"><strong>${label}:</strong> ${value} cm</p>`;
            }
        });
    }

    // Build foto HTML
    let fotoHtml = '';
    if (compressedImageBlobs[laporanIndex]) {
        const fotoSrc = URL.createObjectURL(compressedImageBlobs[laporanIndex]);
        fotoHtml = `<img src="${fotoSrc}" style="max-width: 100%; height: auto; max-height: 280px; border: 1px solid #999; border-radius: 4px; display: block; margin: 5px auto;" />`;
    } else {
        fotoHtml = `<div style="color: #999; font-size: 11px; padding: 80px 10px; text-align: center;">Tidak ada foto</div>`;
    }

    // Tensi HTML (tidak tampil untuk balita)
    let tensiHtml = '';
    if (laporanData.kategori !== 'balita' && laporanData.tensi && laporanData.tensi !== '-') {
        tensiHtml = `<p style="margin: 4px 0;"><strong>Tensi:</strong> ${laporanData.tensi}</p>`;
    }

    page.innerHTML = `
        <div style="margin-bottom: 20px;">
            <h1 style="text-align: center; font-size: 16px; margin: 0 0 5px 0; font-weight: bold; color: #333;">
                LAPORAN KUNJUNGAN RUMAH
            </h1>
            <h1 style="text-align: center; font-size: 16px; margin: 0 0 5px 0; font-weight: bold; color: #333;">
                PUSKESMAS BALAS KLUMPRIK
            </h1>
            <h2 style="text-align: center; font-size: 14px; margin: 0 0 15px 0; font-weight: bold; color: #333;">
                BULAN ${bulan} TAHUN ${tahun}
            </h2>
        </div>
        
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000; font-size: 11px;">
            <thead>
                <tr>
                    <th style="border: 2px solid #000; padding: 10px; background: #e0e0e0; font-weight: bold; text-align: center; width: 8%;">
                        NO
                    </th>
                    <th style="border: 2px solid #000; padding: 10px; background: #e0e0e0; font-weight: bold; text-align: center; width: 18%;">
                        TANGGAL
                    </th>
                    <th style="border: 2px solid #000; padding: 10px; background: #e0e0e0; font-weight: bold; text-align: center; width: 50%;">
                        HASIL
                    </th>
                    <th style="border: 2px solid #000; padding: 10px; background: #e0e0e0; font-weight: bold; text-align: center; width: 24%;">
                        FOTO GEOTAG
                    </th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="border: 2px solid #000; padding: 10px; text-align: center; vertical-align: top;">
                        ${laporanIndex}
                    </td>
                    <td style="border: 2px solid #000; padding: 10px; text-align: center; vertical-align: top;">
                        ${formattedDate}
                    </td>
                    <td style="border: 2px solid #000; padding: 10px; vertical-align: top;">
                        <div style="font-size: 11px; line-height: 1.6;">
                            <p style="margin: 4px 0;"><strong>Nama:</strong> ${laporanData.nama}</p>
                            <p style="margin: 4px 0;"><strong>Usia:</strong> ${laporanData.usia}</p>
                            <p style="margin: 4px 0;"><strong>Alamat:</strong> ${laporanData.alamat}</p>
                            <p style="margin: 4px 0;"><strong>BB:</strong> ${laporanData.beratBadan} Kg</p>
                            <p style="margin: 4px 0;"><strong>TB:</strong> ${laporanData.tinggiBadan} cm</p>
                            ${measurementsHtml}
                            ${tensiHtml}
                            <p style="margin-top: 10px; margin-bottom: 4px;"><strong>Keluhan / Permasalahan:</strong></p>
                            <p style="margin: 4px 0; white-space: pre-wrap;">${laporanData.keluhan}</p>
                            <p style="margin-top: 10px; margin-bottom: 4px;"><strong>Tindak Lanjut:</strong></p>
                            <p style="margin: 4px 0; white-space: pre-wrap;">${laporanData.tindakLanjut}</p>
                        </div>
                    </td>
                    <td style="border: 2px solid #000; padding: 10px; text-align: center; vertical-align: top;">
                        ${fotoHtml}
                    </td>
                </tr>
            </tbody>
        </table>
        
        <div style="margin-top: 15px; padding-top: 10px; border-top: 2px solid #000; font-size: 11px;">
            <p style="margin: 0;"><strong>Kader yang melakukan kunjungan rumah:</strong> ${laporanData.kader}</p>
        </div>
    `;

    return page;
}

/**
 * Initialize PDF generator button
 */
document.addEventListener('DOMContentLoaded', function() {
    const pdfBtn = document.getElementById('generatePdfBtn');
    if (pdfBtn) {
        pdfBtn.addEventListener('click', function(e) {
            e.preventDefault();
            generatePDF();
        });
    }

    console.log('✅ PDF Generator siap digunakan');
    console.log('📄 Tekan tombol "Generate PDF" untuk membuat PDF dengan tampilan identik preview');
});
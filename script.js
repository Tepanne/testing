document.getElementById('file-upload').addEventListener('change', handleFile, false);
document.getElementById('convert-btn').addEventListener('click', convertToVCF, false);

let excelData = null;
let sheetNames = [];

// Fungsi untuk membaca file Excel
function handleFile(event) {
    const reader = new FileReader();
    const file = event.target.files[0];

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Simpan nama sheet
        sheetNames = workbook.SheetNames;

        // Tambahkan nama sheet ke dropdown
        const sheetSelect = document.getElementById('sheet-select');
        sheetSelect.innerHTML = '';
        sheetNames.forEach((name, index) => {
            const option = document.createElement('option');
            option.value = index;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        // Pilih sheet pertama secara default
        sheetSelect.selectedIndex = 0;

        // Tampilkan isi sheet pertama
        displaySheet(workbook, sheetNames[0]);
        
        // Tambah event listener untuk memilih sheet
        sheetSelect.addEventListener('change', () => {
            displaySheet(workbook, sheetNames[sheetSelect.value]);
        });
    };

    reader.readAsArrayBuffer(file);
}

// Fungsi untuk menampilkan isi sheet dalam bentuk tabel
function displaySheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }); // Menjaga kolom kosong

    const table = document.getElementById('excel-table');
    table.innerHTML = ''; // Hapus konten tabel sebelumnya

    // Menambahkan header tabel
    const headerRow = table.insertRow(-1);
    headerRow.className = 'header-row'; // Tambahkan kelas untuk penanda

    // Tambahkan penanda kolom
    const headerCell = document.createElement('th');
    headerCell.className = 'header-column'; // Kelas untuk kolom header
    headerRow.appendChild(headerCell); // Sel untuk penanda kosong

    excelData[0].forEach((header, index) => {
        const cell = headerRow.insertCell(-1);
        cell.textContent = String.fromCharCode(65 + index); // Konversi index ke huruf (A, B, C, ...)
    });

    // Menambahkan isi tabel
    for (let i = 0; i < excelData.length; i++) {
        const row = table.insertRow(-1); // Tambah baris baru
        const rowHeaderCell = row.insertCell(-1);
        rowHeaderCell.textContent = i + 1; // Menambahkan angka baris (1, 2, 3, ...)
        
        excelData[i].forEach((cellData, index) => {
            const cell = row.insertCell(-1); // Tambah sel baru
            cell.textContent = cellData || ''; // Isi sel dengan data, jika ada
        });
    }
}

// Fungsi untuk konversi ke VCF

function convertToVCF() {
    if (!excelData) {
        alert("Silakan unggah file Excel terlebih dahulu.");
        return;
    }

    const startCell = document.getElementById('start-cell').value;
    const baseContactName = document.getElementById('base-contact-name').value;
    const columnLetter = startCell.match(/[A-Za-z]+/)[0].toUpperCase(); // Ambil huruf kolom
    const startRow = parseInt(startCell.match(/\d+/)[0], 10); // Ambil angka baris

    const colIndex = columnLetter.charCodeAt(0) - 65; // Konversi huruf ke indeks kolom (A=0, B=1, dst.)
    const rowIndex = startRow - 1; // Indeks baris dimulai dari 0

    if (rowIndex >= excelData.length || colIndex >= excelData[0].length) {
        alert("Sel awal tidak valid.");
        return;
    }

    let vcfContent = "";

    for (let i = rowIndex; i < excelData.length; i++) {
        let phone = excelData[i][colIndex]; // Ambil nomor telepon di kolom yang ditentukan
        if (phone) {
            phone = String(phone); // Konversi ke string jika belum berupa string

            // Tambahkan tanda + jika belum ada
            if (!phone.startsWith('+')) {
                phone = '+' + phone;
            }

            const contactName = `${baseContactName} ${i - rowIndex + 1}`; // Nama kontak dengan nomor urut
            vcfContent += `BEGIN:VCARD\nVERSION:3.0\nFN:${contactName}\nTEL:${phone}\nEND:VCARD\n\n`;
        }
    }

    // Ambil nama file dari input
    const fileName = document.getElementById('file-name').value || 'contacts'; // Default: 'contacts' jika kosong

    // Buat file VCF dan download
    const blob = new Blob([vcfContent], { type: "text/vcard" });
    const url = URL.createObjectURL(blob);
    const downloadLink = document.getElementById('download-link');
    downloadLink.href = url;
    downloadLink.download = `${fileName}.vcf`; // Gunakan nama file yang diinput
    downloadLink.style.display = 'block';
}

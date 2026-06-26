// ============================================
// BUKU KAS BENDAHARA - GOOGLE APPS SCRIPT v2
// Auto-create sheets if not exist
// ============================================

const SPREADSHEET_ID = '1li0Z3wcCN8LGPakDA83GWBjMyDgCH4bUBTXhwtYNcKA';

// ============================================
// WEB APP ENTRY POINT
// ============================================
function doGet(e) {
  // Initialize sheets if needed
  initializeSheets();
  
  return HtmlService.createHtmlOutputFromFile('bendahara')
    .setTitle('Buku Kas Bendahara')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// INITIALIZE SHEETS (Auto-create if not exist)
// ============================================
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(s => s.getName());
  
  // Create Transaksi sheet if not exist
  if (!sheetNames.includes('Transaksi')) {
    const transSheet = ss.insertSheet('Transaksi');
    transSheet.appendRow(['ID', 'Tanggal', 'Tipe', 'Kategori', 'Deskripsi', 'Jumlah', 'Saldo']);
    
    // Format header
    const headerRange = transSheet.getRange(1, 1, 1, 7);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#667EEA');
    headerRange.setFontColor('#FFFFFF');
    
    // Add sample data
    transSheet.appendRow([1, new Date(), 'Pemasukan', 'Donasi', 'Donasi Bulan Juni', 2000000, 2000000]);
    transSheet.appendRow([2, new Date(), 'Pengeluaran', 'Operasional', 'Beli ATK', 150000, 1850000]);
    
    // Format number columns
    transSheet.getRange('F:G').setNumberFormat('#,##0');
    transSheet.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  }
  
  // Create Kategori sheet if not exist
  if (!sheetNames.includes('Kategori')) {
    const katSheet = ss.insertSheet('Kategori');
    katSheet.appendRow(['Kategori', 'Tipe']);
    
    // Format header
    const headerRange = katSheet.getRange(1, 1, 1, 2);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#667EEA');
    headerRange.setFontColor('#FFFFFF');
    
    // Add default categories
    const categories = [
      ['Donasi', 'Pemasukan'],
      ['Iuran', 'Pemasukan'],
      ['Sponsor', 'Pemasukan'],
      ['Operasional', 'Pengeluaran'],
      ['Rapat', 'Pengeluaran'],
      ['Acara', 'Pengeluaran'],
      ['ATK', 'Pengeluaran'],
      ['Transport', 'Pengeluaran'],
      ['Lainnya', 'Pengeluaran']
    ];
    
    categories.forEach(cat => katSheet.appendRow(cat));
  }
  
  // Create Anggaran sheet if not exist
  if (!sheetNames.includes('Anggaran')) {
    const anggSheet = ss.insertSheet('Anggaran');
    anggSheet.appendRow(['Kategori', 'Anggaran', 'Realisasi', 'Sisa']);
    
    // Format header
    const headerRange = anggSheet.getRange(1, 1, 1, 4);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#667EEA');
    headerRange.setFontColor('#FFFFFF');
    
    // Add default budgets
    const budgets = [
      ['Operasional', 1000000, 0],
      ['Rapat', 500000, 0],
      ['Acara', 2000000, 0],
      ['ATK', 300000, 0],
      ['Transport', 400000, 0]
    ];
    
    budgets.forEach((b, i) => {
      anggSheet.appendRow([b[0], b[1], b[2], `=B${i+2}-C${i+2}`]);
    });
    
    // Format number columns
    anggSheet.getRange('B:D').setNumberFormat('#,##0');
  }

  // Create Pengaturan sheet if not exist
  if (!sheetNames.includes('Pengaturan')) {
    const pengSheet = ss.insertSheet('Pengaturan');
    pengSheet.appendRow(['Pengaturan', 'Nilai']);
    
    // Format header
    const headerRange = pengSheet.getRange(1, 1, 1, 2);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#667EEA');
    headerRange.setFontColor('#FFFFFF');
    
    // Add default image link
    pengSheet.appendRow(['Logo Dashboard', 'https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?w=800&q=80']);
    
    pengSheet.setColumnWidth(1, 150);
    pengSheet.setColumnWidth(2, 400);
  }
  
  // Set column widths for all sheets (Diperbaiki ke ukuran Pixel standar Google Sheets)
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === 'Transaksi') {
      sheet.setColumnWidth(1, 50);  // ID
      sheet.setColumnWidth(2, 120); // Tanggal
      sheet.setColumnWidth(3, 120); // Tipe
      sheet.setColumnWidth(4, 150); // Kategori
      sheet.setColumnWidth(5, 300); // Deskripsi
      sheet.setColumnWidth(6, 150); // Jumlah
      sheet.setColumnWidth(7, 150); // Saldo
    } else if (name === 'Kategori') {
      sheet.setColumnWidth(1, 150); // Kategori
      sheet.setColumnWidth(2, 120); // Tipe
    } else if (name === 'Anggaran') {
      sheet.setColumnWidth(1, 150); // Kategori
      sheet.setColumnWidth(2, 150); // Anggaran
      sheet.setColumnWidth(3, 150); // Realisasi
      sheet.setColumnWidth(4, 150); // Sisa
    } else if (name === 'Pengaturan') {
      sheet.setColumnWidth(1, 150); // Pengaturan
      sheet.setColumnWidth(2, 400); // Nilai/Link
    }
  });
}

// ============================================
// GET ALL DATA
// ============================================
function getData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Get transaksi
  const transSheet = ss.getSheetByName('Transaksi');
  const transData = transSheet.getDataRange().getValues();
  const transaksi = transData.slice(1).map(row => ({
    id: row[0],
    tanggal: formatDate(row[1]),
    tipe: row[2],
    kategori: row[3],
    deskripsi: row[4],
    jumlah: row[5],
    saldo: row[6]
  }));
  
  // Get kategori
  const katSheet = ss.getSheetByName('Kategori');
  const katData = katSheet.getDataRange().getValues();
  const kategori = {};
  katData.slice(1).forEach(row => {
    if (!kategori[row[1]]) kategori[row[1]] = [];
    kategori[row[1]].push(row[0]);
  });
  
  // Get anggaran
  const anggSheet = ss.getSheetByName('Anggaran');
  const anggData = anggSheet.getDataRange().getValues();
  const anggaran = anggData.slice(1).map(row => ({
    kategori: row[0],
    anggaran: row[1],
    realisasi: row[2] || 0,
    sisa: row[3] || row[1] - (row[2] || 0)
  }));

  // Get pengaturan (Logo)
  let logoUrl = '';
  const pengSheet = ss.getSheetByName('Pengaturan');
  if (pengSheet) {
    const pengData = pengSheet.getDataRange().getValues();
    for(let i = 1; i < pengData.length; i++) {
      if(pengData[i][0] === 'Logo Dashboard') {
        logoUrl = pengData[i][1];
        break;
      }
    }
  }
  
  // Calculate summary
  const pemasukan = transaksi.filter(t => t.tipe === 'Pemasukan').reduce((sum, t) => sum + (Number(t.jumlah) || 0), 0);
  const pengeluaran = transaksi.filter(t => t.tipe === 'Pengeluaran').reduce((sum, t) => sum + (Number(t.jumlah) || 0), 0);
  const saldo = pemasukan - pengeluaran;
  
  return {
    transaksi: transaksi.reverse(),
    kategori: kategori,
    anggaran: anggaran,
    settings: {
      logoUrl: logoUrl
    },
    summary: {
      pemasukan: pemasukan,
      pengeluaran: pengeluaran,
      saldo: saldo
    }
  };
}

// ============================================
// ADD TRANSACTION
// ============================================
function addTransaksi(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  
  // Get last ID
  const lastRow = sheet.getLastRow();
  const lastId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0;
  const newId = Number(lastId) + 1;
  
  // Calculate new saldo
  const allData = getData();
  const currentSaldo = allData.summary.saldo;
  const jumlah = Number(data.jumlah);
  const newSaldo = data.tipe === 'Pemasukan' ? currentSaldo + jumlah : currentSaldo - jumlah;
  
  // Add row
  sheet.appendRow([
    newId,
    new Date(data.tanggal),
    data.tipe,
    data.kategori,
    data.deskripsi,
    jumlah,
    newSaldo
  ]);
  
  // Update realisasi anggaran if pengeluaran
  if (data.tipe === 'Pengeluaran') {
    updateAnggaranRealisasi(data.kategori, jumlah);
  }
  
  return { success: true, message: 'Transaksi berhasil ditambahkan' };
}

// ============================================
// UPDATE TRANSACTION
// ============================================
function updateTransaksi(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  
  // Find row by ID
  const dataRange = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < dataRange.length; i++) {
    if (dataRange[i][0] == data.id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: 'Transaksi tidak ditemukan' };
  }
  
  // Update row
  sheet.getRange(rowIndex, 2).setValue(new Date(data.tanggal));
  sheet.getRange(rowIndex, 3).setValue(data.tipe);
  sheet.getRange(rowIndex, 4).setValue(data.kategori);
  sheet.getRange(rowIndex, 5).setValue(data.deskripsi);
  sheet.getRange(rowIndex, 6).setValue(Number(data.jumlah));
  
  // Recalculate all saldo
  recalculateSaldo();
  
  return { success: true, message: 'Transaksi berhasil diupdate' };
}

// ============================================
// DELETE TRANSACTION
// ============================================
function deleteTransaksi(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  
  // Find row by ID
  const dataRange = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < dataRange.length; i++) {
    if (dataRange[i][0] == id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: 'Transaksi tidak ditemukan' };
  }
  
  // Delete row
  sheet.deleteRow(rowIndex);
  
  // Recalculate all saldo
  recalculateSaldo();
  
  return { success: true, message: 'Transaksi berhasil dihapus' };
}

// ============================================
// RECALCULATE SALDO
// ============================================
function recalculateSaldo() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  
  const data = sheet.getDataRange().getValues();
  let saldo = 0;
  
  for (let i = 1; i < data.length; i++) {
    const tipe = data[i][2];
    const jumlah = Number(data[i][5]) || 0;
    
    if (tipe === 'Pemasukan') {
      saldo += jumlah;
    } else {
      saldo -= jumlah;
    }
    
    sheet.getRange(i + 1, 7).setValue(saldo);
  }
}

// ============================================
// UPDATE ANGGARAN REALISASI
// ============================================
function updateAnggaranRealisasi(kategori, jumlah) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Anggaran');
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === kategori) {
      const currentRealisasi = Number(data[i][2]) || 0;
      sheet.getRange(i + 1, 3).setValue(currentRealisasi + jumlah);
      break;
    }
  }
}

// ============================================
// UPDATE NILAI ANGGARAN (MANUAL DARI MODAL)
// ============================================
function updateAnggaran(kategori, jumlahBaru) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Anggaran');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === kategori) {
      sheet.getRange(i + 1, 2).setValue(Number(jumlahBaru));
      
      // Update juga rumus sisa anggarannya
      const cellSisa = sheet.getRange(i + 1, 4);
      cellSisa.setFormula(`=B${i+1}-C${i+1}`);
      
      return { success: true, message: 'Anggaran berhasil diupdate' };
    }
  }
  return { success: false, message: 'Kategori anggaran tidak ditemukan' };
}

// ============================================
// UTILITY FUNCTIONS
// ============================================
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

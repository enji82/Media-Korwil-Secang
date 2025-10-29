function getMuridPaudKelasUsiaData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Murid Kelas");
    
    if (!sheet) {
      throw new Error("Sheet 'Murid Kelas' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridPaudKelasUsiaData', e);
  }
}

/**
 * Mengambil data keadaan murid PAUD menurut jenis kelamin.
 */
function getMuridPaudJenisKelaminData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Murid JK");
    
    if (!sheet) {
      throw new Error("Sheet 'Murid JK' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridPaudJenisKelaminData', e);
  }
}

/**
 * Mengambil data jumlah murid PAUD bulanan.
 */
function getMuridPaudJumlahBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Murid Bulanan");
    
    if (!sheet) {
      throw new Error("Sheet 'Murid Bulanan' tidak ditemukan.");
    }
    
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridPaudJumlahBulananData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut kelas.
 */
function getMuridSdKelasData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Kelas");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Kelas' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdKelasData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut rombel.
 */
function getMuridSdRombelData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Rombel");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Rombel' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdRombelData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut jenis kelamin.
 */
function getMuridSdJenisKelaminData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel JK");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel JK' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdJenisKelaminData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut agama.
 */
function getMuridSdAgamaData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Agama");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Agama' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdAgamaData', e);
  }
}

/**
 * Mengambil data jumlah murid SD bulanan.
 */
function getMuridSdJumlahBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Bulanan");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Bulanan' tidak ditemukan.");
    }
    
    // PERBAIKAN: Menggunakan getValues() untuk mendapatkan angka murni, bukan teks.
    return sheet.getDataRange().getValues();
  } catch (e) {
    return handleError('getMuridSdJumlahBulananData', e);
  }
}
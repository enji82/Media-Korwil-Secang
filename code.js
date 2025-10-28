/**
 * ===================================================================
 * ======================= 1. KONFIGURASI PUSAT ======================
 * ===================================================================
 * Semua ID Spreadsheet dan Folder disimpan di sini agar mudah dikelola.
 */

const SPREADSHEET_CONFIG = {
  // --- Modul SK Pembagian Tugas ---
  SK_BAGI_TUGAS: { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "SK Tabel Kirim" },
  SK_FORM_RESPONSES: { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "Form Responses 1" },
  
  // --- Modul Laporan Bulanan & Data Murid ---
  LAPBUL_PAUD: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Form Responses 1" },
  LAPBUL_SD: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Input" },
  LAPBUL_GABUNGAN: { id: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU" }, // Untuk Riwayat & Status

  // --- Modul Data PTK ---
  PTK_PAUD_KEADAAN: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Keadaan PTK PAUD" },
  PTK_PAUD_JUMLAH_BULANAN: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Data PAUD Bulanan" }, // <-- TAMBAHKAN BARIS INI
  PTK_PAUD_DB: { id: "1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk", sheet: "PTK PAUD" },
  PTK_SD_KEADAAN: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Keadaan PTK SD" },
  PTK_SD_DB: { id: "1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0" },

  // --- Data Pendukung & Dropdown ---
  DATA_SEKOLAH: { id: "1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs" },
  DROPDOWN_DATA: { id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA" },
  FORM_OPTIONS_DB: { id: "1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU" },

  // --- Data SIABA ---
  SIABA_REKAP: { id: "1x3b-yzZbiqP2XfJNRC3XTbMmRTHLd8eEdUqAlKY3v9U", sheet: "Rekap Script" },
  SIABA_TIDAK_PRESENSI: { id: "1mjXz5l_cqBiiR3x9qJ7BU4yQ3f0ghERT9ph8CC608Zc", sheet: "Rekap Script" },
};

const FOLDER_CONFIG = {
  // --- Modul SK Pembagian Tugas ---
  MAIN_SK: "1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs",

  // --- Modul Laporan Bulanan ---
  LAPBUL_KB: "18CxRT-eledBGRtHW1lFd2AZ8Bub6q5ra",
  LAPBUL_TK: "1WUNz_BSFmcwRVlrG67D2afm9oJ-bVI9H",
  LAPBUL_SD: "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El",
};
/**
 * ===================================================================
 * ===================== 2. FUNGSI INTI APLIKASI =====================
 * ===================================================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * ===================================================================
 * ===================== 3. FUNGSI UTILITAS UMUM =====================
 * ===================================================================
 */

function handleError(functionName, error) {
  Logger.log(`Error di ${functionName}: ${error.message}\nStack: ${error.stack}`);
  throw new Error(error.message);
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function getDataFromSheet(configKey) {
  try {
    const config = SPREADSHEET_CONFIG[configKey];
    if (!config) throw new Error(`Konfigurasi untuk '${configKey}' tidak ditemukan.`);
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error(`Sheet '${config.sheet}' tidak ditemukan.`);
    return sheet.getDataRange().getValues();
    // Selalu gunakan getValues() untuk konsistensi data tanggal
  } catch (e) {
    handleError(`getDataFromSheet: ${configKey}`, e);
  }
}

function getCachedData(key, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached != null) {
    Logger.log("Mengambil data dari Cache: " + key);
    return JSON.parse(cached);
  }
  Logger.log("Mengambil data dari Spreadsheet dan menyimpan ke Cache: " + key);
  const freshData = fetchFunction();
  cache.put(key, JSON.stringify(freshData), 21600); // Simpan selama 6 jam
  return freshData;
}


/**
 * ===================================================================
 * ================= 4. MODUL SK PEMBAGIAN TUGAS =====================
 * ===================================================================
 */

function getMasterSkOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
    const getValuesFromSheet = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      return sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
                  .flat()
                  .filter(value => String(value).trim() !== '');
    };

    return {
      'Nama SD': getValuesFromSheet('Nama SD').sort(),
      'Tahun Ajaran': getValuesFromSheet('Tahun Ajaran').sort().reverse(),
      'Semester': getValuesFromSheet('Semester').sort(),
      'Kriteria SK': getValuesFromSheet('Kriteria SK').sort()
    };
  } catch (e) {
    return handleError('getMasterSkOptions', e);
  }
}

function processManualForm(formData) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);  
    
    const tahunAjaranFolderName = formData.tahunAjaran.replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);

    const semesterFolderName = formData.semester;
    const targetFolder = getOrCreateFolder(tahunAjaranFolder, semesterFolderName);

    const newFilename = `${formData.namaSD} - ${tahunAjaranFolderName} - ${formData.semester} - ${formData.kriteriaSK}.pdf`;
    
    const decodedData = Utilities.base64Decode(formData.fileData.data);
    const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
    const newFile = targetFolder.createFile(blob);
    const fileUrl = newFile.getUrl();
    const newRow = [ new Date(), formData.namaSD, formData.tahunAjaran, formData.semester, formData.nomorSK, new Date(formData.tanggalSK), formData.kriteriaSK, fileUrl ];
    
    sheet.appendRow(newRow);
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 6).setNumberFormat("dd-MM-yyyy");

    return "Dokumen SK berhasil diunggah.";
  } catch (e) {
    return handleError('processManualForm', e);
  }
}

/**
 * [REFACTOR - PERBAIKAN PENGURUTAN FINAL] Mengambil data riwayat pengiriman SK.
 * Menggunakan fungsi parseDate yang andal untuk menangani format tanggal yang tidak seragam.
 */
function getSKRiwayatData() {
  try {
    // Gunakan .getValues() untuk mendapatkan objek Date asli jika memungkinkan
    const sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.id)
                                .getSheetByName(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.sheet);

    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }
    
    const allData = sheet.getDataRange().getValues(); // Menggunakan getValues()
    const originalHeaders = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    // Fungsi bantu yang andal untuk mengubah berbagai format tanggal menjadi objek Date
    const parseDate = (value) => {
        if (!value) return new Date(0); // Tanggal paling awal jika kosong
        // Jika sudah berupa objek Date yang valid, kembalikan langsung
        if (value instanceof Date && !isNaN(value)) return value;
        // Jika berupa string, coba parsing format "dd/MM/yyyy HH:mm:ss"
        if (typeof value === 'string') {
          const parts = value.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
          if (parts) {
              // parts[3]=yyyy, parts[2]=MM, parts[1]=dd
              return new Date(parts[3], parseInt(parts[2], 10) - 1, parts[1], parts[4], parts[5], parts[6]);
          }
        }
        // Jika gagal parsing, anggap sebagai tanggal paling awal
        return new Date(0);
    };

    // Ubah setiap baris menjadi objek
    const structuredRows = dataRows.map(row => {
      const rowObject = {};
      originalHeaders.forEach((header, index) => {
        const cell = row[index];
        // Format tanggal menjadi string SETELAH diolah
        if ((header === 'Tanggal Unggah' || header === 'Update') && cell instanceof Date) {
          rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        } else if (header === 'Tanggal SK' && cell instanceof Date) {
          rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
          rowObject[header] = cell;
        }
      });
      return rowObject;
    });

    // **KUNCI PERBAIKAN ADA DI SINI**
    // Urutkan data berdasarkan "Tanggal Unggah" menggunakan fungsi parseDate yang andal
    structuredRows.sort((a, b) => {
      const dateB = parseDate(b['Tanggal Unggah']);
      const dateA = parseDate(a['Tanggal Unggah']);
      return dateB - dateA; // Urutan descending (terbaru di atas)
    });
    
    // Tentukan urutan header yang akan ditampilkan di tabel
    const desiredHeaders = ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Tanggal SK", "Kriteria SK", "Dokumen", "Tanggal Unggah"];

    return {
      headers: desiredHeaders,
      rows: structuredRows
    };

  } catch (e) {
    return handleError('getSKRiwayatData', e);
  }
}

/**
 * [REFACTOR - PERBAIKAN FINAL] Mengambil data status pengiriman SK.
 * Mengembalikan data dalam format objek agar sesuai dengan fungsi generik.
 * Urutan baris sesuai dengan spreadsheet sumber.
 */
function getSKStatusData() {
  try {
    const data = getDataFromSheet('SK_BAGI_TUGAS');
    if (!data || data.length < 2) {
      return { headers: [], rows: [] };
    }

    const headers = data[0].map(h => String(h).trim());
    const dataRows = data.slice(1);

    // **AWAL BLOK PERBAIKAN**
    // Mengubah setiap baris (yang tadinya array) menjadi objek
    const structuredRows = dataRows.map(row => {
      const rowObject = {};
      headers.forEach((header, index) => {
        rowObject[header] = row[index];
      });
      return rowObject;
    });
    // **AKHIR BLOK PERBAIKAN**

    // Tidak ada pengurutan di sini, sesuai permintaan sebelumnya
    
    return {
      headers: headers,
      rows: structuredRows // Kirim data yang sudah berformat objek
    };

  } catch (e) {
    return handleError('getSKStatusData', e);
  }
}

function getSKKelolaData() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }

    const originalData = sheet.getDataRange().getValues();
    const originalHeaders = originalData[0].map(h => String(h).trim());
    const dataRows = originalData.slice(1);

    // Fungsi bantu untuk mengubah tanggal menjadi objek Date yang valid untuk perbandingan
    const parseDate = (value) => {
      // Jika sudah berupa objek Date, kembalikan langsung
      if (value instanceof Date && !isNaN(value)) return value;
      // Jika tidak ada nilai, kembalikan tanggal paling awal agar diurutkan ke bawah
      if (!value) return new Date(0); 
      // Coba parsing dari format string
      const date = new Date(value);
      return isNaN(date) ? new Date(0) : date;
    };

    const indexedData = dataRows.map((row, index) => ({
      row: row,
      originalIndex: index + 2
    }));

    // **KUNCI PERBAIKAN 1: LOGIKA PENGURUTAN MULTI-LEVEL**
    const updateIndex = originalHeaders.indexOf('Update');
    const timestampIndex = originalHeaders.indexOf('Tanggal Unggah');

    indexedData.sort((a, b) => {
      // Prioritas 1: Urutkan berdasarkan kolom "Update" (terbaru di atas)
      const dateB_update = parseDate(b.row[updateIndex]);
      const dateA_update = parseDate(a.row[updateIndex]);
      if (dateB_update.getTime() !== dateA_update.getTime()) {
        return dateB_update - dateA_update;
      }

      // Prioritas 2: Jika "Update" sama, urutkan berdasarkan "Tanggal Unggah" (terbaru di atas)
      const dateB_timestamp = parseDate(b.row[timestampIndex]);
      const dateA_timestamp = parseDate(a.row[timestampIndex]);
      return dateB_timestamp - dateA_timestamp;
    });

    // Mengubah baris menjadi objek (logika ini tetap sama)
    const structuredRows = indexedData.map(item => {
      const rowObject = {
        _rowIndex: item.originalIndex,
        _source: 'SK'
      };
      originalHeaders.forEach((header, i) => {
        let cell = item.row[i];
        if ((header === 'Tanggal Unggah' || header === 'Update' || header === 'Tanggal SK') && cell instanceof Date) {
          const format = (header === 'Tanggal SK') ? "dd/MM/yyyy" : "dd/MM/yyyy HH:mm:ss";
          rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), format);
        } else {
          rowObject[header] = cell;
        }
      });
      return rowObject;
    });
    
    // **KUNCI PERBAIKAN 2: MENAMBAHKAN KOLOM BARU**
    // Tambahkan "Tanggal Unggah" dan "Update" setelah "Aksi"
    const desiredHeaders = ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Kriteria SK", "Dokumen", "Aksi", "Tanggal Unggah", "Update"];

    return {
      headers: desiredHeaders,
      rows: structuredRows
    };
  } catch (e) {
    return handleError("getSKKelolaData", e);
  }
}

function getSKDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header] = rowValues[i];
    });
    return rowData;
  } catch (e) {
    return handleError("getSKDataByRow", e);
  }
}

function updateSKData(formData) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const range = sheet.getRange(formData.rowIndex, 1, 1, headers.length);
    const existingRowValues = range.getDisplayValues()[0]; // getDisplayValues untuk perbandingan string
    const existingRowObject = {};
    headers.forEach((header, i) => { existingRowObject[header] = existingRowValues[i]; });

    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
    const tahunAjaranFolderName = existingRowObject['Tahun Ajaran'].replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);
    
    let fileUrl = existingRowObject['Dokumen'];
    const fileUrlIndex = headers.indexOf('Dokumen');

    // Tentukan nama file dan folder tujuan BARU berdasarkan data form
    const newSemesterFolderName = formData['Semester'];
    const newTargetFolder = getOrCreateFolder(tahunAjaranFolder, newSemesterFolderName);
    const newFilename = `${existingRowObject['Nama SD']} - ${tahunAjaranFolderName} - ${newSemesterFolderName} - ${formData['Kriteria SK']}.pdf`;

    // Skenario 1: File BARU diunggah
    if (formData.fileData && formData.fileData.data) {
      // Hapus file lama jika ada
      if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        try {
          const fileId = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
          if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch (e) {
          Logger.log(`Gagal menghapus file lama saat upload baru: ${e.message}`);
        }
      }
      
      // Unggah file baru
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
      const newFile = newTargetFolder.createFile(blob);
      fileUrl = newFile.getUrl();

    // Skenario 2: TIDAK ada file baru, tapi data (Semester/Kriteria) berubah
    } else if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        const fileIdMatch = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
        if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const file = DriveApp.getFileById(fileId);
            const currentFileName = file.getName();
            const currentParentFolder = file.getParents().next();

            // Cek jika nama file atau folder perlu diubah
            if (currentFileName !== newFilename || currentParentFolder.getName() !== newSemesterFolderName) {
                file.moveTo(newTargetFolder);
                file.setName(newFilename);
                fileUrl = file.getUrl(); // Perbarui URL setelah dipindah/diubah namanya
                Logger.log(`File dipindahkan ke folder '${newSemesterFolderName}' dan diubah namanya menjadi '${newFilename}'`);
            }
        }
    }
    
    formData['Dokumen'] = fileUrl; // Pastikan URL file diperbarui di formData
    formData['Update'] = new Date();

    const newRowValuesForSheet = headers.map(header => {
      return formData.hasOwnProperty(header) ? formData[header] : existingRowObject[header];
    });

    // Gunakan getRange(row, col, numRows, numCols) dan setValues()
    sheet.getRange(formData.rowIndex, 1, 1, headers.length).setValues([newRowValuesForSheet]);
    
    const tanggalSKIndex = headers.indexOf('Tanggal SK');
    if (tanggalSKIndex !== -1) {
      sheet.getRange(formData.rowIndex, tanggalSKIndex + 1).setNumberFormat("dd-MM-yyyy");
    }
    
    return "Data berhasil diperbarui!";
  } catch (e) {
    return handleError('updateSKData', e);
  }
}


function deleteSKData(rowIndex, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fileUrlIndex = headers.findIndex(h => h.trim().toLowerCase() === 'dokumen');
    
    if (fileUrlIndex !== -1) {
        const fileUrl = sheet.getRange(rowIndex, fileUrlIndex + 1).getValue();
        if (fileUrl && typeof fileUrl === 'string') {
            const fileId = fileUrl.match(/[-\w]{25,}/);
            if (fileId) {
                try {
                    DriveApp.getFileById(fileId[0]).setTrashed(true);
                } catch (err) {
                    Logger.log(`Gagal menghapus file dengan ID ${fileId[0]}: ${err.message}`);
                }
            }
        }
    }
    
    sheet.deleteRow(rowIndex);
    return "Data dan file terkait berhasil dihapus.";
  } catch (e) {
    return handleError("deleteSKData", e);
  }
}


/**
 * ===================================================================
 * ==================== 5. MODUL LAPORAN BULAN =======================
 * ===================================================================
 */

/**
 * [REFACTOR - PERBAIKAN] Mengambil daftar sekolah PAUD.
 * Menggunakan kunci SPREADSHEET_CONFIG yang sudah diperbarui.
 */
function getPaudSchoolLists() {
  const cacheKey = 'paud_school_lists';
  return getCachedData(cacheKey, function() {
    // **PERBAIKAN DI BARIS INI**
    // Menggunakan DROPDOWN_DATA, bukan DATA_SEKOLAH_PAUD
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id); 
    
    const getValues = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().filter(Boolean).sort();
    };
    return { KB: getValues('Nama KB'), TK: getValues('Nama TK') };
  });
}

function getSdSchoolLists() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
    const getValues = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().filter(Boolean).sort();
    };
    return {
      SDN: getValues('Nama SDN'),
      SDS: getValues('Nama SDS')
    };
  } catch (e) {
    return handleError('getSdSchoolLists', e);
  }
}

function processLapbulFormPaud(formData) {
  try {
    const jenjang = formData.jenjang;
    let FOLDER_ID_LAPBUL;
    if (jenjang === 'KB') {
      FOLDER_ID_LAPBUL = FOLDER_CONFIG.LAPBUL_KB;
    } else if (jenjang === 'TK') {
      FOLDER_ID_LAPBUL = FOLDER_CONFIG.LAPBUL_TK;
    } else {
      throw new Error("Jenjang tidak valid: " + jenjang);
    }
    
    const mainFolder = DriveApp.getFolderById(FOLDER_ID_LAPBUL);
    const tahunFolder = getOrCreateFolder(mainFolder, formData.tahun);
    const bulanFolder = getOrCreateFolder(tahunFolder, formData.laporanBulan);

    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    const newFileName = `${formData.namaSekolah} - Lapbul ${formData.laporanBulan} ${formData.tahun}.pdf`;
    const newFile = bulanFolder.createFile(blob).setName(newFileName);
    const fileUrl = newFile.getUrl();

    const config = SPREADSHEET_CONFIG.LAPBUL_PAUD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);

    const newRow = [
      new Date(),
      formData.laporanBulan, formData.tahun, formData.npsn, formData.statusSekolah, formData.jumlahRombel, formData.jenjang, formData.namaSekolah,
      formData.murid_0_1_L, formData.murid_0_1_P, formData.murid_1_2_L, formData.murid_1_2_P,
      formData.murid_2_3_L, formData.murid_2_3_P, formData.murid_3_4_L, formData.murid_3_4_P,
      formData.murid_4_5_L, formData.murid_4_5_P, formData.murid_5_6_L, formData.murid_5_6_P,
      formData.murid_6_up_L, formData.murid_6_up_P, formData.kelompok_A_L, formData.kelompok_A_P,
      formData.kelompok_B_L, formData.kelompok_B_P, formData.kepsek_ASN, formData.kepsek_Non_ASN,
      formData.guru_PNS, formData.guru_PPPK, formData.guru_GTY, formData.guru_GTT,
      formData.tendik_Penjaga, formData.tendik_TAS, formData.tendik_Pustakawan, formData.tendik_Lainnya,
      fileUrl
    ];
    sheet.appendRow(newRow);

    return "Sukses! Laporan Bulan PAUD berhasil dikirim.";
  } catch (e) {  
    return handleError('processLapbulFormPaud', e);
  }
}

function processLapbulFormSd(formData) {
  try {
    // 1. Proses File dan Simpan ke Google Drive
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.LAPBUL_SD);
    const tahunFolder = getOrCreateFolder(mainFolder, formData.tahun);
    const bulanFolder = getOrCreateFolder(tahunFolder, formData.laporanBulan);

    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    
    const newFileName = `${formData.namaSekolah} - Lapbul ${formData.laporanBulan} ${formData.tahun}.pdf`;
    const newFile = bulanFolder.createFile(blob).setName(newFileName);
    const fileUrl = newFile.getUrl();

    // 2. Siapkan Data untuk Disimpan ke Google Sheet
    const config = SPREADSHEET_CONFIG.LAPBUL_SD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const getValue = (key) => formData[key] || 0;

    // Susun baris baru sesuai urutan kolom A s/d GH
    const newRowLengkap = [
      new Date(), // A
      formData.laporanBulan, // B
      formData.tahun, // C
      formData.statusSekolah, // D
      formData.namaSekolah, // E
      formData.npsn, // F
      formData.jumlahRombel, // G
      fileUrl, // H
      
      // Kelas 1 (I-AC)
      getValue('k1_jumlah_rombel'), getValue('k1_rombel_tunggal_L'), getValue('k1_rombel_tunggal_P'),
      getValue('k1_rombel_a_L'), getValue('k1_rombel_a_P'), getValue('k1_rombel_b_L'), getValue('k1_rombel_b_P'), getValue('k1_rombel_c_L'), getValue('k1_rombel_c_P'),
      getValue('k1_agama_islam_L'), getValue('k1_agama_islam_P'), getValue('k1_agama_kristen_L'), getValue('k1_agama_kristen_P'), getValue('k1_agama_katolik_L'), getValue('k1_agama_katolik_P'), 
      getValue('k1_agama_hindu_L'), getValue('k1_agama_hindu_P'), getValue('k1_agama_buddha_L'), getValue('k1_agama_buddha_P'), getValue('k1_agama_konghucu_L'), getValue('k1_agama_konghucu_P'),
      
      // Kelas 2 (AD-AX)
      getValue('k2_jumlah_rombel'), getValue('k2_rombel_tunggal_L'), getValue('k2_rombel_tunggal_P'),
      getValue('k2_rombel_a_L'), getValue('k2_rombel_a_P'), getValue('k2_rombel_b_L'), getValue('k2_rombel_b_P'), getValue('k2_rombel_c_L'), getValue('k2_rombel_c_P'),
      getValue('k2_agama_islam_L'), getValue('k2_agama_islam_P'), getValue('k2_agama_kristen_L'), getValue('k2_agama_kristen_P'), getValue('k2_agama_katolik_L'), getValue('k2_agama_katolik_P'), 
      getValue('k2_agama_hindu_L'), getValue('k2_agama_hindu_P'), getValue('k2_agama_buddha_L'), getValue('k2_agama_buddha_P'), getValue('k2_agama_konghucu_L'), getValue('k2_agama_konghucu_P'),

    
      // Kelas 3 (AY-BS)
      getValue('k3_jumlah_rombel'), getValue('k3_rombel_tunggal_L'), getValue('k3_rombel_tunggal_P'),
      getValue('k3_rombel_a_L'), getValue('k3_rombel_a_P'), getValue('k3_rombel_b_L'), getValue('k3_rombel_b_P'), getValue('k3_rombel_c_L'), getValue('k3_rombel_c_P'),
      getValue('k3_agama_islam_L'), getValue('k3_agama_islam_P'), getValue('k3_agama_kristen_L'), getValue('k3_agama_kristen_P'), getValue('k3_agama_katolik_L'), getValue('k3_agama_katolik_P'), 
      getValue('k3_agama_hindu_L'), getValue('k3_agama_hindu_P'), getValue('k3_agama_buddha_L'), getValue('k3_agama_buddha_P'), getValue('k3_agama_konghucu_L'), getValue('k3_agama_konghucu_P'),
      
      // Kelas 4 (BT-CN)
      getValue('k4_jumlah_rombel'), getValue('k4_rombel_tunggal_L'), getValue('k4_rombel_tunggal_P'),
      getValue('k4_rombel_a_L'), getValue('k4_rombel_a_P'), getValue('k4_rombel_b_L'), getValue('k4_rombel_b_P'), getValue('k4_rombel_c_L'), getValue('k4_rombel_c_P'),
      getValue('k4_agama_islam_L'), getValue('k4_agama_islam_P'), getValue('k4_agama_kristen_L'), getValue('k4_agama_kristen_P'), getValue('k4_agama_katolik_L'), getValue('k4_agama_katolik_P'), 
      getValue('k4_agama_hindu_L'), 
      getValue('k4_agama_hindu_P'), getValue('k4_agama_buddha_L'), getValue('k4_agama_buddha_P'), getValue('k4_agama_konghucu_L'), getValue('k4_agama_konghucu_P'),

      // Kelas 5 (CO-DI)
      getValue('k5_jumlah_rombel'), getValue('k5_rombel_tunggal_L'), getValue('k5_rombel_tunggal_P'),
      getValue('k5_rombel_a_L'), getValue('k5_rombel_a_P'), getValue('k5_rombel_b_L'), getValue('k5_rombel_b_P'), getValue('k5_rombel_c_L'), getValue('k5_rombel_c_P'),
      getValue('k5_agama_islam_L'), getValue('k5_agama_islam_P'), getValue('k5_agama_kristen_L'), getValue('k5_agama_kristen_P'), getValue('k5_agama_katolik_L'), getValue('k5_agama_katolik_P'), 
      getValue('k5_agama_hindu_L'), getValue('k5_agama_hindu_P'), getValue('k5_agama_buddha_L'), getValue('k5_agama_buddha_P'), getValue('k5_agama_konghucu_L'), getValue('k5_agama_konghucu_P'),

      // Kelas 6 (DJ-ED)
      getValue('k6_jumlah_rombel'), getValue('k6_rombel_tunggal_L'), getValue('k6_rombel_tunggal_P'),
      getValue('k6_rombel_a_L'), getValue('k6_rombel_a_P'), getValue('k6_rombel_b_L'), getValue('k6_rombel_b_P'), getValue('k6_rombel_c_L'), getValue('k6_rombel_c_P'),
      getValue('k6_agama_islam_L'), getValue('k6_agama_islam_P'), getValue('k6_agama_kristen_L'), getValue('k6_agama_kristen_P'), getValue('k6_agama_katolik_L'), getValue('k6_agama_katolik_P'), 
     
      getValue('k6_agama_hindu_L'), getValue('k6_agama_hindu_P'), getValue('k6_agama_buddha_L'), getValue('k6_agama_buddha_P'), getValue('k6_agama_konghucu_L'), getValue('k6_agama_konghucu_P'),

      // PTK Negeri (EE-FH)
      getValue('ptk_negeri_ks_pns'), getValue('ptk_negeri_ks_pppk'),
      getValue('ptk_negeri_guru_kelas_pns'), getValue('ptk_negeri_guru_kelas_pppk'), getValue('ptk_negeri_guru_kelas_pppkpw'), getValue('ptk_negeri_guru_kelas_gtt'),
      getValue('ptk_negeri_guru_pai_pns'), getValue('ptk_negeri_guru_pai_pppk'), getValue('ptk_negeri_guru_pai_pppkpw'), getValue('ptk_negeri_guru_pai_gtt'),
      getValue('ptk_negeri_guru_pjok_pns'), getValue('ptk_negeri_guru_pjok_pppk'), getValue('ptk_negeri_guru_pjok_pppkpw'), getValue('ptk_negeri_guru_pjok_gtt'),
      getValue('ptk_negeri_guru_kristen_pns'), getValue('ptk_negeri_guru_kristen_pppk'), getValue('ptk_negeri_guru_kristen_pppkpw'), getValue('ptk_negeri_guru_kristen_gtt'),
      getValue('ptk_negeri_guru_inggris_gtt'), getValue('ptk_negeri_guru_lainnya_gtt'),
      getValue('ptk_negeri_tendik_operator_pppk'), getValue('ptk_negeri_tendik_operator_pppkpw'), getValue('ptk_negeri_tendik_operator_ptt'),
      getValue('ptk_negeri_tendik_pengelola_pppk'), getValue('ptk_negeri_tendik_pengelola_pppkpw'), getValue('ptk_negeri_tendik_pengelola_ptt'),
      getValue('ptk_negeri_tendik_penjaga_ptt'), getValue('ptk_negeri_tendik_tas_ptt'), getValue('ptk_negeri_tendik_pustakawan_ptt'), getValue('ptk_negeri_tendik_lainnya_ptt'),

      // PTK Swasta (FI-GH)
 
      getValue('ptk_swasta_ks_gty'), getValue('ptk_swasta_ks_gtt'),
      getValue('ptk_swasta_guru_kelas_gty'), getValue('ptk_swasta_guru_kelas_gtt'),
      getValue('ptk_swasta_guru_pai_gty'), getValue('ptk_swasta_guru_pai_gtt'),
      getValue('ptk_swasta_guru_pjok_gty'), getValue('ptk_swasta_guru_pjok_gtt'),
      getValue('ptk_swasta_guru_kristen_gty'), getValue('ptk_swasta_guru_kristen_gtt'),
      getValue('ptk_swasta_guru_inggris_gty'), getValue('ptk_swasta_guru_inggris_gtt'),
      getValue('ptk_swasta_guru_lainnya_gty'), getValue('ptk_swasta_guru_lainnya_gtt'),
      getValue('ptk_swasta_tendik_operator_pty'), getValue('ptk_swasta_tendik_operator_ptt'),
      getValue('ptk_swasta_tendik_pengelola_pty'), getValue('ptk_swasta_tendik_pengelola_ptt'),
      getValue('ptk_swasta_tendik_penjaga_pty'), getValue('ptk_swasta_tendik_penjaga_ptt'),
      getValue('ptk_swasta_tendik_tas_pty'), getValue('ptk_swasta_tendik_tas_ptt'),
      getValue('ptk_swasta_tendik_pustakawan_pty'), getValue('ptk_swasta_tendik_pustakawan_ptt'),
      getValue('ptk_swasta_tendik_lainnya_pty'), getValue('ptk_swasta_tendik_lainnya_ptt')
    ];
    sheet.appendRow(newRowLengkap);

    // 3. Simpan ringkasan data LANGSUNG ke Sheet "Riwayat"
    const configRiwayat = SPREADSHEET_CONFIG.LAPBUL_GABUNGAN;
    const sheetRiwayat = SpreadsheetApp.openById(configRiwayat.id).getSheetByName("Riwayat");
    
    // Sesuaikan urutan kolom dengan sheet "Riwayat"
    const newRowRiwayat = [
      new Date(),              // Kolom A: Tanggal Unggah
      formData.namaSekolah,    // Kolom B: Nama Sekolah
      "SD",                    // Kolom C: Jenjang
      formData.statusSekolah,  // Kolom D: Status
      formData.laporanBulan,   // Kolom E: Bulan
      formData.tahun,          // Kolom F: Tahun
      formData.jumlahRombel,   // Kolom G: Rombel
      fileUrl                  // Kolom H: Dokumen
    ];
    sheetRiwayat.appendRow(newRowRiwayat);
    // **AKHIR BLOK KODE BARU**

    // 4. Pastikan semua perubahan disimpan sebelum skrip selesai
    SpreadsheetApp.flush();
    return "Sukses! Laporan Bulan SD berhasil dikirim.";
  } catch (e) {
    return handleError('processLapbulFormSd', e);
  }
}

/**
 * [REFACTOR] Mengambil dan menggabungkan data riwayat dari sumber SD dan PAUD secara real-time.
 * Ini adalah solusi untuk masalah delay IMPORTRANGE.
 */
function getLapbulRiwayatData() {
  try {
    const desiredHeaders = ["Nama Sekolah", "Jenjang", "Bulan", "Tahun", "Rombel", "Dokumen", "Tanggal Unggah"];
    let combinedData = [];

    // Fungsi bantu untuk memproses data mentah dari setiap sheet
    const processSheetData = (sheet, mapping) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      
      const data = sheet.getDataRange().getValues(); // Gunakan getValues untuk tanggal
      const dataRows = data.slice(1);

      dataRows.forEach(row => {
        // Lewati baris jika kolom Tanggal Unggah (kolom A di kedua sumber) kosong
        if (!row[0]) return;

        const rowObject = {
          "Nama Sekolah":   row[mapping.namaSekolah],
          "Jenjang":        mapping.jenjang === 'static' ? 'SD' : row[mapping.jenjang],
          "Bulan":          row[mapping.bulan],
          "Tahun":          row[mapping.tahun],
          "Rombel":         row[mapping.rombel],
          "Dokumen":        row[mapping.dokumen],
          "Tanggal Unggah": row[mapping.tanggalUnggah]
        };
        combinedData.push(rowObject);
      });
    };
    
    // 1. Proses Data SD
    const sdConfig = SPREADSHEET_CONFIG.LAPBUL_SD;
    const sdSheet = SpreadsheetApp.openById(sdConfig.id).getSheetByName(sdConfig.sheet);
    const sdMapping = { namaSekolah: 4, jenjang: 'static', bulan: 1, tahun: 2, rombel: 6, dokumen: 7, tanggalUnggah: 0 }; // E=4, B=1, C=2, G=6, H=7, A=0
    processSheetData(sdSheet, sdMapping);

    // 2. Proses Data PAUD
    const paudConfig = SPREADSHEET_CONFIG.LAPBUL_PAUD;
    const paudSheet = SpreadsheetApp.openById(paudConfig.id).getSheetByName(paudConfig.sheet);
    const paudMapping = { namaSekolah: 7, jenjang: 6, bulan: 1, tahun: 2, rombel: 5, dokumen: 36, tanggalUnggah: 0 }; // H=7, G=6, B=1, C=2, F=5, AK=36, A=0
    processSheetData(paudSheet, paudMapping);
    
    // 3. Urutkan semua data yang sudah digabung
    const parseDate = (value) => {
        if (value instanceof Date && !isNaN(value)) return value;
        return new Date(0);
    };
    combinedData.sort((a, b) => {
      const dateB = parseDate(b["Tanggal Unggah"]);
      const dateA = parseDate(a["Tanggal Unggah"]);
      return dateB - dateA;
    });

    // 4. Format tanggal menjadi string setelah diurutkan
    const formattedData = combinedData.map(row => {
      row["Tanggal Unggah"] = (row["Tanggal Unggah"] instanceof Date) 
        ? Utilities.formatDate(row["Tanggal Unggah"], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
        : row["Tanggal Unggah"];
      return row;
    });

    return {
      headers: desiredHeaders,
      rows: formattedData
    };

  } catch (e) {
    return handleError('getLapbulRiwayatData', e);
  }
}

function getLapbulStatusData() {
  try {
    const sheet = SpreadsheetApp.openById("1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU").getSheetByName("Status");
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }
    
    const allData = sheet.getDataRange().getDisplayValues();
    const sourceHeaders = allData[0];
    const dataRows = allData.slice(1);

    // Tentukan kolom mana yang ingin ditampilkan dari spreadsheet sumber
    const displayIndices = [2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
    const finalHeaders = displayIndices.map(index => sourceHeaders[index]);

    // Buat data baris yang sudah dipilah
    const finalRows = dataRows.map(row => {
      // Simpan data filter secara terpisah untuk diolah di klien
      const rowObject = {
        _filterJenjang: row[0], // Kolom A
        _filterTahun: row[1],   // Kolom B
        _filterStatus: row[3]  // Kolom D
      };
      // Isi data utama untuk ditampilkan
      finalHeaders.forEach((header, index) => {
        rowObject[header] = row[displayIndices[index]];
      });
      return rowObject;
    });

    return {
      headers: finalHeaders,
      rows: finalRows
    };

  } catch (e) {
    return handleError('getLapbulStatusData', e);
  }
}

/**
 * [REFACTOR - PERBAIKAN PENGURUTAN] Mengambil data gabungan Laporan Bulan.
 * Mengurutkan berdasarkan prioritas: 1. Update (terbaru), 2. Tanggal Unggah (terbaru).
 * Mampu menangani format tanggal yang tidak seragam.
 */
function getLapbulKelolaData() {
  try {
    const paudConfig = SPREADSHEET_CONFIG.LAPBUL_PAUD;
    const sdConfig = SPREADSHEET_CONFIG.LAPBUL_SD;

    const paudSheet = SpreadsheetApp.openById(paudConfig.id).getSheetByName(paudConfig.sheet);
    const sdSheet = SpreadsheetApp.openById(sdConfig.id).getSheetByName(sdConfig.sheet);
    
    let combinedData = [];

    // **FUNGSI BANTU BARU YANG LEBIH ANDAL UNTUK MEMBACA TANGGAL**
    const parseDate = (value) => {
        // Jika tidak ada nilai, kembalikan tanggal paling awal agar diurutkan ke bawah
        if (!value) return new Date(0); 
        // Jika sudah berupa objek Date yang valid, kembalikan langsung
        if (value instanceof Date && !isNaN(value)) return value;
        // Jika berupa string, coba parsing format "dd/MM/yyyy HH:mm:ss"
        if (typeof value === 'string') {
          const parts = value.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
          if (parts) {
              // parts[3]=yyyy, parts[2]=MM, parts[1]=dd
              return new Date(parts[3], parseInt(parts[2], 10) - 1, parts[1], parts[4], parts[5], parts[6]);
          }
        }
        // Jika gagal parsing, anggap sebagai tanggal paling awal
        return new Date(0);
    };

    const processSheetData = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      // Gunakan .getValues() untuk mendapatkan objek Date asli jika ada
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());
      const rows = data.slice(1);

      rows.forEach((row, index) => {
        if (!row[0]) return;
        const rowObject = {
          _rowIndex: index + 2,
          _source: sourceName
        };
        headers.forEach((header, i) => {
          // Simpan objek Date asli untuk pengurutan
          rowObject[header] = row[i];
        });
        if (sourceName === 'SD') {
          rowObject['Jenjang'] = 'SD';
        }
        combinedData.push(rowObject);
      });
    };

    processSheetData(paudSheet, 'PAUD');
    processSheetData(sdSheet, 'SD');
    
    // **BLOK PENGURUTAN MULTI-LEVEL YANG SUDAH DIPERBAIKI**
    combinedData.sort((a, b) => {
        // Prioritas 1: Urutkan berdasarkan kolom "Update" (terbaru di atas)
        const dateB_update = parseDate(b['Update']);
        const dateA_update = parseDate(a['Update']);
        if (dateB_update.getTime() !== dateA_update.getTime()) {
            return dateB_update - dateA_update;
        }

        // Prioritas 2: Jika "Update" sama, urutkan berdasarkan "Tanggal Unggah" (terbaru di atas)
        const dateB_timestamp = parseDate(b['Tanggal Unggah']);
        const dateA_timestamp = parseDate(a['Tanggal Unggah']);
        return dateB_timestamp - dateA_timestamp;
    });
    
    // Format tanggal menjadi string SETELAH diurutkan
    const formattedRows = combinedData.map(row => {
        if (row['Update'] instanceof Date) {
            row['Update'] = Utilities.formatDate(row['Update'], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        }
        if (row['Tanggal Unggah'] instanceof Date) {
            row['Tanggal Unggah'] = Utilities.formatDate(row['Tanggal Unggah'], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        }
        return row;
    });

    const finalHeaders = ["Nama Sekolah", "Jenjang", "Bulan", "Tahun", "Dokumen", "Aksi", "Tanggal Unggah", "Update"];

    return { headers: finalHeaders, rows: formattedRows };

  } catch (e) {
    return handleError('getLapbulKelolaData', e);
  }
}

/**
 * [REFACTOR - PERBAIKAN KONSISTENSI v2.0] Mengambil satu baris data Laporan Bulan.
 * Mengubah nama header 'Jumlah Rombel' menjadi 'Rombel' untuk PAUD dan SD.
 */
function getLapbulDataByRow(rowIndex, source) {
  try {
    let sheet;
    if (source === 'PAUD') {
      const config = SPREADSHEET_CONFIG.LAPBUL_PAUD;
      sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    } else if (source === 'SD') {
      const config = SPREADSHEET_CONFIG.LAPBUL_SD;
      sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const rowData = {};

    headers.forEach((header, i) => {
      let key = header.trim();
      // **PERBAIKAN KONSISTENSI DI SINI**
      if (key === 'Jumlah Rombel') {
        key = 'Rombel'; // Ganti nama header agar cocok dengan form
      }
      rowData[key] = values[i];
    });
    
    return rowData;
  } catch (e) {
    return handleError('getLapbulDataByRow', e);
  }
}


/**
 * [REFACTOR - PERBAIKAN] Memperbarui satu baris data Laporan Bulan.
 * Menggunakan kunci SPREADSHEET_CONFIG yang benar.
 */
function updateLapbulData(formData) {
  try {
    let sheet, FOLDER_ID;
    const source = formData.source;
    const rowIndex = formData.rowIndex;
    if (!source || !rowIndex) throw new Error("Informasi 'source' atau 'rowIndex' tidak ditemukan.");

    // **AWAL BLOK PERBAIKAN**
    if (source === 'PAUD') {
      const config = SPREADSHEET_CONFIG.LAPBUL_PAUD;
      sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
      const jenjang = formData.jenjang || sheet.getRange(rowIndex, headers.indexOf('Jenjang') + 1).getValue();
      FOLDER_ID = jenjang === 'KB' ? FOLDER_CONFIG.LAPBUL_KB : FOLDER_CONFIG.LAPBUL_TK;
    } else if (source === 'SD') {
      const config = SPREADSHEET_CONFIG.LAPBUL_SD;
      sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
      FOLDER_ID = FOLDER_CONFIG.LAPBUL_SD;
    // **AKHIR BLOK PERBAIKAN**
    } else {
      throw new Error("Sumber data tidak valid: " + source);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());

    if (!formData.Bulan || !formData.Tahun) {
        throw new Error("Informasi 'Bulan' atau 'Tahun' kosong. Tidak dapat memproses file.");
    }

    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const existingValues = range.getValues()[0];

    if (formData.fileData && formData.fileData.data) {
      const docHeaderName = 'Dokumen';
      const docIndex = headers.indexOf(docHeaderName);
      if (docIndex !== -1 && existingValues[docIndex]) {
        try {
          const fileId = String(existingValues[docIndex]).match(/[-\w]{25,}/);
          if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch (e) { Logger.log(`Gagal menghapus file lama: ${e.message}`); }
      }
      
      const mainFolder = DriveApp.getFolderById(FOLDER_ID);
      const tahunFolder = getOrCreateFolder(mainFolder, formData.Tahun);
      const bulanFolder = getOrCreateFolder(tahunFolder, formData.Bulan);
      
      const namaSekolah = formData['Nama Sekolah'] || existingValues[headers.indexOf('Nama Sekolah')];
      const newFileName = `${namaSekolah} - Lapbul ${formData.Bulan} ${formData.Tahun}.pdf`;
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFileName);
      const newFile = bulanFolder.createFile(blob);
      formData[docHeaderName] = newFile.getUrl();
    }

    formData['Update'] = new Date();
    const newRowValues = headers.map((header, index) => {
      // Ganti nama 'Rombel' dari form menjadi 'Jumlah Rombel' untuk disimpan di sheet
      if (header === 'Jumlah Rombel' && formData.hasOwnProperty('Rombel')) {
        return formData['Rombel'];
      }
      if (formData.hasOwnProperty(header)) {
        return formData[header];
      }
      return existingValues[index];
    });
    range.setValues([newRowValues]);

    const updateColIndex = headers.indexOf('Update');
    if (updateColIndex !== -1) {
      sheet.getRange(rowIndex, updateColIndex + 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
    
    return "Data berhasil diperbarui.";
  } catch (e) {
    Logger.log(`Error in updateLapbulData: ${e.message}\nStack: ${e.stack}`);
    return { error: `Terjadi error di server: ${e.message}` };
  }
}

/**
 * Menghapus data Laporan Bulan dari sheet dan file terkait dari Drive.
 * @param {number} rowIndex Nomor baris di spreadsheet.
 * @param {string} source Sumber data ('PAUD' atau 'SD').
 * @param {string} deleteCode Kode hapus (yyyyMMdd).
 * @returns {string} Pesan sukses.
 */
function deleteLapbulData(rowIndex, source, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    let sheet;
    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const docIndex = headers.findIndex(h => h.trim().toLowerCase() === 'dokumen');
    if (docIndex !== -1) {
        const fileUrl = sheet.getRange(rowIndex, docIndex + 1).getValue();
        if (fileUrl && typeof fileUrl === 'string') {
            const fileId = fileUrl.match(/[-\w]{25,}/);
            if (fileId) {
                try {
                    DriveApp.getFileById(fileId[0]).setTrashed(true);
                } catch (err) {
                    Logger.log(`Gagal menghapus file Laporan Bulan dengan ID ${fileId[0]}: ${err.message}`);
                }
            }
        }
    }
    
    sheet.deleteRow(rowIndex);
    return "Data Laporan Bulan dan file terkait berhasil dihapus.";
  } catch (e) {
    return handleError("deleteLapbulData", e);
  }
}

/**
 * ===================================================================
 * =================== 6. MODUL GOOGLE DRIVE (ARSIP) =================
 * ===================================================================
 */

function getFolders(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const subFolders = parentFolder.getFolders();
    const folderList = [];
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      folderList.push({
        id: folder.getId(),
        name: folder.getName()
      });
    }
    folderList.sort((a, b) => b.name.localeCompare(a.name));
    return folderList;
  } catch (e) {
    return handleError("getFolders", e);
  }
}

function getFiles(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const files = parentFolder.getFiles();
    const fileList = [];
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        url: file.getUrl()
      });
    }
    fileList.sort((a, b) => a.name.localeCompare(b.name));
    return fileList;
  } catch (e) {
    return handleError("getFiles", e);
  }
}

function getLapbulInfo() {
  const cacheKey = 'lapbul_info_v1'; // Kunci unik untuk cache
  return getCachedData(cacheKey, function() {
    try {
      // Menggunakan kembali konfigurasi DROPDOWN_DATA karena ID spreadsheet sama
      const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
      const sheet = ss.getSheetByName('Informasi');
      
      if (!sheet) {
        throw new Error("Sheet 'Informasi' tidak ditemukan di spreadsheet referensi.");
      }

      const lastRow = sheet.getLastRow();
 
      if (lastRow < 2) return []; // Kembalikan array kosong jika tidak ada data

      // Ambil data dari A2 sampai baris terakhir yang berisi konten
      const range = sheet.getRange('A2:A' + lastRow);
      const values = range.getValues()
                          .flat() // Mengubah array 2D menjadi 1D
              
              .filter(item => String(item).trim() !== ''); // Menghapus baris kosong
      return values;
    } catch (e) {
      return handleError('getLapbulInfo', e);
    }
  });
}

function getUnduhFormatInfo() {
  const cacheKey = 'unduh_format_info_v1';
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
      const sheet = ss.getSheetByName('Informasi');
      
      if (!sheet) {
        throw new Error("Sheet 'Informasi' tidak ditemukan.");
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];

      // Ambil data dari B2 sampai baris 
      const range = sheet.getRange('B2:B' + lastRow);
      const values = range.getDisplayValues()
                          .flat()
                          .filter(item => String(item).trim() !== '');
      return values;
    } catch (e) {
      return handleError('getUnduhFormatInfo', e);
  
    }
  });
}

/**
 * [BARU] Mengambil data keadaan PTK PAUD dari sheet 'Jumlah PTK Bulanan'.
 * Mengembalikan data dalam format objek yang siap diolah oleh fungsi generik.
 */
function getPtkKeadaanPaudData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_KEADAAN;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    // Mengubah setiap baris menjadi objek
    const structuredRows = dataRows.map(row => {
      const rowObject = {};
      headers.forEach((header, index) => {
        rowObject[header] = row[index];
      });
      // Menambahkan data tersembunyi untuk filter berdasarkan kolom "Jenjang" (Kolom B)
      rowObject._filterJenjang = row[headers.indexOf('Jenjang')];
      return rowObject;
    });

    return {
      headers: headers,
      rows: structuredRows
    };
  } catch (e) {
    return handleError('getPtkKeadaanPaudData', e);
  }
}

function getKeadaanPtkSdData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("Keadaan PTK SD");
    if (!sheet) {
      throw new Error("Sheet 'Keadaan PTK SD' tidak ditemukan.");
    }
    
    return sheet.getDataRange().getDisplayValues();

  } catch (e) {
    return handleError('getKeadaanPtkSdData', e);
  }
}

/**
 * [REFACTOR] Mengambil data jumlah PTK bulanan untuk PAUD.
 * Menggunakan SPREADSHEET_CONFIG terpusat.
 */
function getPtkJumlahBulananPaudData() {
  try {
    // **PERBAIKAN DI SINI: Menggunakan konfigurasi terpusat**
    const config = SPREADSHEET_CONFIG.PTK_PAUD_JUMLAH_BULANAN;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }
    
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const desiredHeaders = [
        "Nama Lembaga", "Jenjang", "Status", "Jumlah Rombel", 
        "KS PNS", "KS Non PNS", "GK PNS", "GK PPPK", "GTY", "GTT", 
        "Penjaga", "TAS", "Pustakawan", "Tendik Lain", "Jumlah PTK"
    ];
    
    const colIndices = desiredHeaders.map(header => headers.indexOf(header));
    const tahunIndex = headers.indexOf('Tahun');
    const bulanIndex = headers.indexOf('Bulan');

    const structuredRows = dataRows.map(row => {
        const rowObject = {
            _filterTahun: row[tahunIndex],
            _filterBulan: row[bulanIndex],
            _filterJenjang: row[headers.indexOf('Jenjang')]
        };
        desiredHeaders.forEach((header, i) => {
            const index = colIndices[i];
            rowObject[header] = row[index];
        });
        return rowObject;
    });

    return {
      headers: desiredHeaders,
      rows: structuredRows
    };

  } catch (e) {
    return handleError('getPtkJumlahBulananPaudData', e);
  }
}

/**
 * [REFACTOR] Mengambil data daftar PTK PAUD.
 * Mengembalikan data dalam format objek yang sudah diurutkan berdasarkan Nama.
 */
function getDaftarPtkPaudData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    // Mengurutkan baris data berdasarkan abjad pada kolom "Nama" (indeks ke-2)
    const namaIndex = headers.indexOf('Nama');
    if (namaIndex > -1) {
      dataRows.sort((a, b) => {
        return a[namaIndex].localeCompare(b[namaIndex]);
      });
    }
    
    // Tentukan kolom mana saja yang ingin ditampilkan
    const desiredHeaders = ["Nama", "L/P", "Jabatan", "TMT", "Pendidikan", "Serdik", "NUPTK"];
    
    // Ubah setiap baris menjadi objek
    const structuredRows = dataRows.map(row => {
      const rowObject = {
        // Data untuk filter
        _filterJenjang: row[headers.indexOf('Jenjang')],
        _filterNamaLembaga: row[headers.indexOf('Nama Lembaga')],
        _filterStatus: row[headers.indexOf('Status')]
      };
      // Data untuk ditampilkan
      desiredHeaders.forEach(header => {
        const index = headers.indexOf(header);
        if (index > -1) {
          rowObject[header] = row[index];
        }
      });
      return rowObject;
    });

    return {
      headers: desiredHeaders,
      rows: structuredRows
    };

  } catch (e) {
    return handleError('getDaftarPtkPaudData', e);
  }
}

/**
 * [REFACTOR - PERBAIKAN PENGURUTAN KOMBINASI] Mengambil data PTK PAUD untuk halaman Kelola.
 * Mengurutkan berdasarkan tanggal terbaru dari kolom "Update" atau "Tanggal Input".
 */
function getKelolaPtkPaudData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }

    const allData = sheet.getDataRange().getValues();
    const headers = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    const parseDate = (value) => {
      if (value instanceof Date && !isNaN(value)) return value;
      return new Date(0); // Return tanggal paling awal jika tidak valid
    };

    const indexedData = dataRows.map((row, index) => ({
      originalRowIndex: index + 2,
      rowData: row
    }));

    // **AWAL BLOK KODE PERBAIKAN PENGURUTAN**
    const updateIndex = headers.indexOf('Update');
    const dateInputIndex = headers.indexOf('Tanggal Input');

    indexedData.sort((a, b) => {
      // Ambil tanggal terbaru dari baris B (update atau input)
      const dateB_update = parseDate(b.rowData[updateIndex]);
      const dateB_input = parseDate(b.rowData[dateInputIndex]);
      const latestDateB = dateB_update > dateB_input ? dateB_update : dateB_input;

      // Ambil tanggal terbaru dari baris A (update atau input)
      const dateA_update = parseDate(a.rowData[updateIndex]);
      const dateA_input = parseDate(a.rowData[dateInputIndex]);
      const latestDateA = dateA_update > dateA_input ? dateA_update : dateA_input;
      
      // Bandingkan tanggal terbaru dari kedua baris
      return latestDateB - latestDateA;
    });
    // **AKHIR BLOK KODE PERBAIKAN PENGURUTAN**

    const structuredRows = indexedData.map(item => {
      const rowObject = {
        _rowIndex: item.originalRowIndex,
        _source: 'PAUD'
      };
      headers.forEach((header, i) => {
        let cell = item.rowData[i];
        if (cell instanceof Date) {
          rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        } else {
          rowObject[header] = cell || "";
        }
      });
      return rowObject;
    });
    
    const desiredHeaders = ['Nama', 'Nama Lembaga', 'Status', 'NIP/NIY', 'Jabatan', 'Aksi', 'Tanggal Input', 'Update'];

    return { headers: desiredHeaders, rows: structuredRows };
  } catch (e) {
    return handleError('getKelolaPtkPaudData', e);
  }
}

/**
 * Mengambil data satu baris PTK PAUD berdasarkan nomor barisnya.
 */
function getPtkPaudDataByRow(rowIndex) {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      // Untuk tanggal, ambil nilai mentahnya agar bisa di-format di form
      if (header === 'TMT' && values[i] instanceof Date) {
        rowData[header] = Utilities.formatDate(values[i], "UTC", "yyyy-MM-dd");
      } else {
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    Logger.log(`Error in getPtkPaudDataByRow: ${e.message}`);
    return { error: e.message };
  }
}

function updatePtkPaudData(formData) {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const rowIndex = formData.rowIndex;
    if (!rowIndex) throw new Error("Nomor baris (rowIndex) tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const oldValues = range.getValues()[0]; // Ambil nilai yang sudah ada

    // Buat objek dari nilai lama untuk memudahkan akses
    const oldRowData = {};
    headers.forEach((header, i) => {
      oldRowData[header] = oldValues[i];
    });

    // Tambahkan timestamp update
    formData['Update'] = new Date();

    // Buat baris baru dengan mempertahankan nilai lama jika tidak ada di form
    const newRowValues = headers.map(header => {
      // [KUNCI PERBAIKAN] Jika header ada di form, gunakan nilai baru.
      // Jika tidak (seperti 'Tanggal Input'), gunakan nilai lama.
      if (formData.hasOwnProperty(header)) {
        return formData[header];
      } else {
        return oldRowData[header];
      }
    });

    range.setValues([newRowValues]);

    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    Logger.log(`Error in updatePtkPaudData: ${e.message}`);
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

function getNewPtkPaudOptions() {
  try {
    // ID Spreadsheet baru untuk sumber data form
    const ss = SpreadsheetApp.openById("1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU");
    const sheet = ss.getSheetByName("Form PAUD");
    if (!sheet) throw new Error("Sheet 'Form PAUD' tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // Ambil kolom A sampai D

    const jenjangOptions = [...new Set(data.map(row => row[0]).filter(Boolean))].sort();
    const statusOptions = [...new Set(data.map(row => row[2]).filter(Boolean))].sort();
    const jabatanOptions = [...new Set(data.map(row => row[3]).filter(Boolean))].sort();

    // Buat pemetaan Nama Lembaga berdasarkan Jenjang
    const lembagaMap = {};
    data.forEach(row => {
      const jenjang = row[0];
      const lembaga = row[1];
      if (jenjang && lembaga) {
        if (!lembagaMap[jenjang]) {
          lembagaMap[jenjang] = [];
        }
        if (!lembagaMap[jenjang].includes(lembaga)) {
          lembagaMap[jenjang].push(lembaga);
        }
      }
    });
    // Urutkan setiap daftar lembaga
    for (const jenjang in lembagaMap) {
      lembagaMap[jenjang].sort();
    }

    return {
      'Jenjang': jenjangOptions,
      'Nama Lembaga': lembagaMap,
      'Status': statusOptions,
      'Jabatan': jabatanOptions
    };
  } catch (e) {
    Logger.log(`Error in getNewPtkPaudOptions: ${e.message}`);
    return { error: e.message };
  }
}

/**
 * Menambahkan data PTK PAUD baru ke spreadsheet.
 */
function addNewPtkPaud(formData) {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(header => {
      if (header === 'Tanggal Input') {
        return new Date();
      }
      return formData[header] || ""; // Gunakan string kosong jika tidak ada data
    });

    // appendRow akan selalu menambahkan data di baris paling bawah
    sheet.appendRow(newRow);
    
    // [PERBAIKAN] Dua baris kode untuk mengurutkan data telah dihapus dari sini

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    Logger.log(`Error in addNewPtkPaud: ${e.message}`);
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

/**
 * Menghapus data PTK PAUD dari spreadsheet.
 */
function deletePtkPaudData(rowIndex, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");
    
    // Validasi rowIndex untuk memastikan itu angka dan dalam rentang yang valid
    const maxRows = sheet.getLastRow();
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > maxRows) {
      throw new Error("Nomor baris tidak valid atau di luar jangkauan.");
    }

    sheet.deleteRow(rowIndex);
    
    return "Data PTK berhasil dihapus.";
  } catch (e) {
    Logger.log(`Error in deletePtkPaudData: ${e.message}`);
    throw new Error(`Gagal menghapus data: ${e.message}`);
  }
}

/**
 * [REFACTOR - PERBAIKAN FINAL] Mengambil data keadaan PTK SD.
 * Mengembalikan data baris sebagai array mentah untuk menangani header duplikat,
 * dan mengirimkan indeks kolom filter secara terpisah.
 */
function getPtkKeadaanSdData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_SD_KEADAAN;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 3) { // Butuh minimal 3 baris
      return { headerRow1: [], headerRow2: [], rows: [], filterConfigs: [] };
    }

    const allData = sheet.getDataRange().getDisplayValues();
    const headerRow1 = allData[0];
    const headerRow2 = allData[1];
    const dataRows = allData.slice(2);

    // Siapkan konfigurasi filter untuk dikirim ke klien
    const filterConfigs = [];
    const statusIndex = headerRow1.indexOf('Status');
    if (statusIndex > -1) {
      filterConfigs.push({
        id: 'filterStatusSekolah',
        columnIndex: statusIndex // Kirim INDEKS kolom, bukan nama
      });
    }

    return {
      isDoubleHeader: true, // Tambahkan flag untuk penanda
      headerRow1: headerRow1,
      headerRow2: headerRow2,
      rows: dataRows, // Kirim sebagai array mentah
      filterConfigs: filterConfigs // Kirim konfigurasi filter
    };
  } catch (e) {
    return handleError('getPtkKeadaanSdData', e);
  }
}

function getJumlahPtkSdBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("PTK Bulanan SD");
    if (!sheet) {
      throw new Error("Sheet 'PTK Bulanan SD' tidak ditemukan.");
    }
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getJumlahPtkSdBulananData', e);
  }
}

/**
 * Mengambil data daftar PTK SD Negeri dari spreadsheet.
 */
function getDaftarPtkSdnData() {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    const sheet = ss.getSheetByName("PTK SDN");
    if (!sheet) {
      throw new Error("Sheet 'PTK SDN' tidak ditemukan.");
    }
    // Mengirim seluruh data mentah, pemrosesan dilakukan di client-side
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdnData', e);
  }
}

/**
 * Mengambil data daftar PTK SD Swasta dari spreadsheet.
 */
function getDaftarPtkSdsData() {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    const sheet = ss.getSheetByName("PTK SDS");
    if (!sheet) {
      throw new Error("Sheet 'PTK SDS' tidak ditemukan.");
    }
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdsData', e);
  }
}

/**
 * Mengambil data gabungan PTK SDN dan SDS untuk halaman Kelola.
 */
function getKelolaPtkSdData() {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    const sdnSheet = ss.getSheetByName("PTK SDN");
    const sdsSheet = ss.getSheetByName("PTK SDS");

    let combinedData = [];

    const processSheet = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      const data = sheet.getDataRange().getDisplayValues(); // Menggunakan DisplayValues lebih aman
      const rows = data.slice(1);

      const idx = {
          unitKerja: 0, nama: 1, status: 3, nipNiy: 4, jabatan: 9, tglInput: 12, update: 13
      };

      rows.forEach((row, index) => {
        if (!row[idx.nama]) return; // Lewati baris jika nama kosong
        
        combinedData.push({
          rowIndex: index + 2,
          source: sourceName,
          data: {
            'Unit Kerja': row[idx.unitKerja],
            'Nama': row[idx.nama],
            'Status': row[idx.status],
            'NIP/NIY': row[idx.nipNiy],
            'Jabatan': row[idx.jabatan],
            'Tanggal Input': row[idx.tglInput],
            'Update': row[idx.update]
          }
        });
      });
    };

    processSheet(sdnSheet, 'SDN');
    processSheet(sdsSheet, 'SDS');
    
    // TIDAK ADA PENGURUTAN DI SINI, HANYA MENGIRIM DATA
    return { rows: combinedData };

  } catch (e) {
    Logger.log(`Error in getKelolaPtkSdData: ${e.message}`);
    return { error: e.message };
  }
}

/**
 * Mengambil opsi-opsi untuk form Tambah PTK SD.
 */
function getNewPtkSdOptions() {
  try {
    const dataSekolahSS = SpreadsheetApp.openById("1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs");
    const getValues = (sheetName) => {
      const sheet = dataSekolahSS.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().filter(Boolean).sort();
    };

    return {
      'Nama SDN': getValues('Nama SDN'),
      'Nama SDS': getValues('Nama SDS'),
      'Status Kepegawaian': getValues('Status Kepegawaian SD'),
      'Pangkat': getValues('Pangkat'),
      'Jabatan': getValues('Jabatan')
    };
  } catch (e) {
    return handleError('getNewPtkSdOptions', e);
  }
}

/**
 * Menambahkan data PTK SD baru ke spreadsheet yang sesuai.
 */
function addNewPtkSd(formData) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    
    if (formData.statusSekolah === 'Negeri') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (formData.statusSekolah === 'Swasta') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Status Sekolah tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet untuk status '${formData.statusSekolah}' tidak ditemukan.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    // Tambahkan timestamp
    formData['Tanggal Input'] = new Date();

    const newRow = headers.map(header => formData[header] || ""); // Gunakan string kosong jika tidak ada data

    sheet.appendRow(newRow);

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    Logger.log(`Error in addNewPtkSd: ${e.message}`);
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

/**
 * Mengambil semua opsi dropdown untuk form Tambah/Edit PTK SD.
 */
function getNewPtkSdOptions() {
  try {
    const ss = SpreadsheetApp.openById("1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU");
    const sheet = ss.getSheetByName("Form SD");
    if (!sheet) throw new Error("Sheet 'Form SD' tidak ditemukan.");

    const getUniqueValues = (col) => {
      const data = sheet.getRange(`${col}2:${col}${sheet.getLastRow()}`).getDisplayValues().flat().filter(Boolean);
      return [...new Set(data)].sort();
    };

    return {
      'Unit Kerja Negeri': getUniqueValues('B'),
      'Unit Kerja Swasta': getUniqueValues('A'),
      'Status Negeri': getUniqueValues('C'),
      'Pangkat PNS': getUniqueValues('D'),
      'Pangkat PPPK': getUniqueValues('E'),
      'Pangkat PPPK PW': getUniqueValues('F'),
      'Jabatan': getUniqueValues('G'),
      'Tugas Tambahan': getUniqueValues('H'),
      'Status Swasta': getUniqueValues('I'),
    };
  } catch (e) {
    return handleError('getNewPtkSdOptions', e);
  }
}

/**
 * Menambahkan data PTK SD baru ke sheet yang sesuai.
 */
function addNewPtkSd(formData) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    
    if (formData.statusSekolah === 'Negeri') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (formData.statusSekolah === 'Swasta') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Status Sekolah tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet untuk status '${formData.statusSekolah}' tidak ditemukan.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    formData['Input'] = new Date();

    const newRow = headers.map(header => {
      let value = formData[header] || "";
      // KUNCI PERBAIKAN 2: Jika header adalah NUPTK dan ada nilainya, tambahkan apostrof
      if (header === 'NUPTK' && value) {
        return "'" + value;
      }
      return value;
    });

    sheet.appendRow(newRow);

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    Logger.log(`Error in addNewPtkSd: ${e.message}`);
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

/**
 * Mengambil data satu baris PTK SD berdasarkan nomor baris dan sumbernya.
 */
function getPtkSdDataByRow(rowIndex, source) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid.");
    }
    if (!sheet) throw new Error(`Sheet '${source}' tidak ditemukan.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      // Untuk tanggal, ambil nilai mentahnya agar bisa di-format di form
      if ((header === 'TMT' || header === 'TMT CPNS' || header === 'TMT PNS') && values[i] instanceof Date) {
        rowData[header] = Utilities.formatDate(values[i], "UTC", "yyyy-MM-dd");
      } else {
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    return handleError('getPtkSdDataByRow', e);
  }
}

/**
 * Memperbarui data PTK SD yang ada di spreadsheet.
 */
function updatePtkSdData(formData) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    const source = formData.source;
    
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet '${source}' tidak ditemukan.`);
    const rowIndex = formData.rowIndex;
    if (!rowIndex) throw new Error("Nomor baris (rowIndex) tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const oldValues = range.getValues()[0];

    // Tambahkan timestamp update
    formData['Update'] = new Date();
    
    // NUPTK sebagai teks
    if (formData['NUPTK']) {
      formData['NUPTK'] = "'" + formData['NUPTK'];
    }

    const newRowValues = headers.map((header, index) => {
      // Jika header ada di form, gunakan nilai baru. Jika tidak, gunakan nilai lama.
      return formData.hasOwnProperty(header) ? formData[header] : oldValues[index];
    });

    range.setValues([newRowValues]);

    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    Logger.log(`Error in updatePtkSdData: ${e.message}`);
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

/**
 * Mengambil data Kebutuhan PTK SD Negeri.
 */
function getKebutuhanPtkSdnData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("Kebutuhan Guru");
    if (!sheet) {
      throw new Error("Sheet 'Kebutuhan Guru' tidak ditemukan.");
    }
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getKebutuhanPtkSdnData', e);
  }
}

/**
 * Menghapus data PTK SD (Negeri atau Swasta) dari spreadsheet.
 */
function deletePtkSdData(rowIndex, source, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid: " + source);
    }

    if (!sheet) throw new Error("Sheet sumber '" + source + "' tidak ditemukan.");
    
    const maxRows = sheet.getLastRow();
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > maxRows) {
      throw new Error("Nomor baris tidak valid atau di luar jangkauan.");
    }

    sheet.deleteRow(rowIndex);
    
    return "Data PTK berhasil dihapus.";
  } catch (e) {
    Logger.log(`Error in deletePtkSdData: ${e.message}`);
    throw new Error(`Gagal menghapus data: ${e.message}`);
  }
}

/**
 * Mengambil data keadaan murid PAUD menurut kelas dan usia.
 */
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

/**
 * ===================================================================
 * ======================= 7. MODUL SIABA ============================
 * ===================================================================
 */

/**
 * Mengambil opsi filter (Tahun, Bulan, Unit Kerja) untuk Daftar Presensi SIABA.
 */
function getSiabaFilterOptions() {
  try {
    // Mengambil data Unit Kerja dari sheet "Unit Siaba"
    const ssDropdown = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
    const sheetUnitKerja = ssDropdown.getSheetByName("Unit Siaba");
    let unitKerjaOptions = [];
    if (sheetUnitKerja && sheetUnitKerja.getLastRow() > 1) {
      unitKerjaOptions = sheetUnitKerja.getRange(2, 1, sheetUnitKerja.getLastRow() - 1, 1)
                                      .getDisplayValues()
                                      .flat()
                                      .filter(Boolean)
                                      .sort();
    }

    // Mengambil data Tahun dan Bulan dari Spreadsheet SIABA
    const ssSiaba = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SIABA_REKAP.id);
    const sheetSiaba = ssSiaba.getSheetByName(SPREADSHEET_CONFIG.SIABA_REKAP.sheet);
    if (!sheetSiaba || sheetSiaba.getLastRow() < 2) {
         throw new Error("Sheet Rekap SIABA tidak ditemukan atau kosong.");
    }

    const tahunBulanData = sheetSiaba.getRange(2, 1, sheetSiaba.getLastRow() - 1, 2).getDisplayValues();
    const uniqueTahun = [...new Set(tahunBulanData.map(row => row[0]))].filter(Boolean).sort().reverse();
    const uniqueBulan = [...new Set(tahunBulanData.map(row => row[1]))].filter(Boolean);
    
    const monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    uniqueBulan.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    return {
      'Tahun': uniqueTahun,
      'Bulan': uniqueBulan,
      'Unit Kerja': unitKerjaOptions
    };
  } catch (e) {
    return handleError('getSiabaFilterOptions', e);
  }
}


/**
 * Mengambil data presensi SIABA berdasarkan filter TAHUN dan BULAN.
 * FINAL: Mengurutkan berdasarkan 6 level prioritas.
 * Prioritas: TP -> TA -> PLA -> TAp -> TU -> Nama.
 */
function getSiabaPresensiData(filters) {
  try {
    const { tahun, bulan } = filters;
    if (!tahun || !bulan) {
      throw new Error("Filter Tahun dan Bulan wajib diisi.");
    }

    const config = SPREADSHEET_CONFIG.SIABA_REKAP;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
        return { headers: [], rows: [] };
    }

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => {
      const tahunMatch = String(row[0]) === String(tahun); // Kolom A
      const bulanMatch = String(row[1]) === String(bulan); // Kolom B
      return tahunMatch && bulanMatch;
    });

    const startIndex = 2;
    const endIndex = 86;

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => {
       return row.slice(startIndex, endIndex + 1);
    });
    
    // --- BLOK PENGURUTAN MULTI-LEVEL (FINAL) ---
    
    const tpIndex = displayHeaders.indexOf('TP');
    const taIndex = displayHeaders.indexOf('TA');
    const plaIndex = displayHeaders.indexOf('PLA');
    const tapIndex = displayHeaders.indexOf('TAp');
    const tuIndex = displayHeaders.indexOf('TU'); // <-- Tambahkan Prioritas 5
    const namaIndex = displayHeaders.indexOf('Nama');

    // Lakukan pengurutan hanya jika kolom prioritas utama (TP) ditemukan
    if (tpIndex !== -1) {
      displayRows.sort((a, b) => {
        const compareDesc = (index) => {
            if (index === -1) return 0; // Jika kolom tidak ditemukan, jangan urutkan
            const valB = parseInt(b[index], 10) || 0;
            const valA = parseInt(a[index], 10) || 0;
            return valB - valA;
        };
        
        // Prioritas 1: Urutkan berdasarkan TP
        let diff = compareDesc(tpIndex);
        if (diff !== 0) return diff;

        // Prioritas 2: Jika TP sama, urutkan berdasarkan TA
        diff = compareDesc(taIndex);
        if (diff !== 0) return diff;

        // Prioritas 3: Jika TA sama, urutkan berdasarkan PLA
        diff = compareDesc(plaIndex);
        if (diff !== 0) return diff;

        // Prioritas 4: Jika PLA sama, urutkan berdasarkan TAp
        diff = compareDesc(tapIndex);
        if (diff !== 0) return diff;
        
        // Prioritas 5: Jika TAp sama, urutkan berdasarkan TU
        diff = compareDesc(tuIndex);
        if (diff !== 0) return diff;

        // Prioritas 6 (Fallback): Jika semua sama, urutkan berdasarkan Nama
        if (namaIndex !== -1) {
            const namaA = a[namaIndex] || "";
            const namaB = b[namaIndex] || "";
            return namaA.localeCompare(namaB);
        }

        return 0;
      });
    }
    // --- AKHIR BLOK PENGURUTAN ---

    return { headers: displayHeaders, rows: displayRows };

  } catch (e) {
    return handleError('getSiabaPresensiData', e);
  }
}

/**
 * Mengambil opsi filter (Tahun, Bulan, Unit Kerja) dari sheet "ASN Tidak Presensi".
 */
function getSiabaTidakPresensiFilterOptions() {
  try {
    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
         throw new Error("Sheet 'Rekap Script' untuk data Tidak Presensi tidak ditemukan atau kosong.");
    }

    const filterData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues(); // Kolom A, B, C
    const uniqueTahun = [...new Set(filterData.map(row => row[0]))].filter(Boolean).sort().reverse();
    const uniqueBulan = [...new Set(filterData.map(row => row[1]))].filter(Boolean);
    const uniqueUnitKerja = [...new Set(filterData.map(row => row[2]))].filter(Boolean).sort();
    
    const monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    uniqueBulan.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    return {
      'Tahun': uniqueTahun,
      'Bulan': uniqueBulan,
      'Unit Kerja': uniqueUnitKerja
    };
  } catch (e) {
    return handleError('getSiabaTidakPresensiFilterOptions', e);
  }
}

/**
 * Mengambil data ASN Tidak Presensi berdasarkan filter.
 * Diurutkan berdasarkan 'Jumlah' (terbanyak) lalu 'Nama' (abjad).
 */
function getSiabaTidakPresensiData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    if (!tahun || !bulan) {
      throw new Error("Filter Tahun dan Bulan wajib diisi.");
    }

    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
        return { headers: [], rows: [] };
    }

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => {
      const tahunMatch = String(row[0]) === String(tahun);
      const bulanMatch = String(row[1]) === String(bulan);
      const unitKerjaMatch = (unitKerja === "Semua") || (String(row[2]) === String(unitKerja));
      return tahunMatch && bulanMatch && unitKerjaMatch;
    });

    const startIndex = 3; // Kolom D
    const endIndex = 8;   // Kolom I

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => {
       return row.slice(startIndex, endIndex + 1);
    });
    
    // --- BLOK PENGURUTAN BARU ---
    // Cari posisi (indeks) kolom 'Jumlah' dan 'Nama' secara dinamis
    const jumlahIndex = displayHeaders.indexOf('Jumlah'); // Seharusnya di posisi ke-4
    const namaIndex = displayHeaders.indexOf('Nama');     // Seharusnya di posisi ke-0

    // Lakukan pengurutan
    displayRows.sort((a, b) => {
      // Prioritas 1: Urutkan berdasarkan Jumlah (menurun)
      const valB_jumlah = (jumlahIndex !== -1) ? (parseInt(b[jumlahIndex], 10) || 0) : 0;
      const valA_jumlah = (jumlahIndex !== -1) ? (parseInt(a[jumlahIndex], 10) || 0) : 0;
      if (valB_jumlah !== valA_jumlah) {
        return valB_jumlah - valA_jumlah;
      }

      // Prioritas 2 (Fallback): Jika Jumlah sama, urutkan berdasarkan Nama (abjad A-Z)
      if (namaIndex !== -1) {
          const namaA = a[namaIndex] || "";
          const namaB = b[namaIndex] || "";
          return namaA.localeCompare(namaB);
      }
      return 0;
    });
    // --- AKHIR BLOK PENGURUTAN ---

    return { headers: displayHeaders, rows: displayRows };
  } catch (e) {
    return handleError('getSiabaTidakPresensiData', e);
  }
}
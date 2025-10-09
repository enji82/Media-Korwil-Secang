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
  
  // --- Modul Laporan Bulanan ---
  LAPBUL_FORM_RESPONSES_PAUD: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Form Responses 1" },
  LAPBUL_FORM_RESPONSES_SD: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Input" },
  LAPBUL_RIWAYAT: { id: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU", sheet: "Riwayat" },
  LAPBUL_STATUS: { id: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU", sheet: "Status" },

  // --- Data Pendukung ---
  DATA_SEKOLAH: { id: "1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs", sheet: "Data Sekolah" },
  DATA_SEKOLAH_PAUD: { id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA" }, 
  DROPDOWN_DATA: { id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA" },
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
    return sheet.getDataRange().getValues(); // Selalu gunakan getValues() untuk konsistensi data tanggal
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

function getRiwayatPengirimanSKData() {
  try {
    const data = getDataFromSheet('SK_FORM_RESPONSES');
    if (data.length < 2) return data;
    
    const headers = data[0].map(h => String(h).trim());
    let dataRows = data.slice(1);
    
    const timestampIndex = headers.indexOf('Tanggal Unggah');
    const tglSKIndex = headers.indexOf('Tanggal SK');
    
    const sortIndex = (timestampIndex > -1) ? timestampIndex : 0;

    const parseDate = (value) => {
        if (!value) return new Date(0);
        if (value instanceof Date && !isNaN(value)) return value;
        const date = new Date(value);
        return isNaN(date) ? new Date(0) : date;
    };
    
    dataRows.sort((a, b) => {
      const dateA = parseDate(a[sortIndex]);
      const dateB = parseDate(b[sortIndex]);
      return dateB - dateA;
    });
    
    const formattedDataRows = dataRows.map(row => {
        return row.map((cell, index) => {
            if (cell instanceof Date) {
                if (index === tglSKIndex) {
                    return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
                } else {
                    return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
                }
            }
            return cell;
        });
    });
    
    return [headers].concat(formattedDataRows);

  } catch(e) {
    return handleError('getRiwayatPengirimanSKData', e);
  }
}

function getStatusPengirimanSKData() { 
  const data = getDataFromSheet('SK_BAGI_TUGAS');
  return data.map(row => row.map(cell => String(cell))); // Pastikan semua data adalah string
}

function getSKDataForManagement() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error(`Sheet '${config.sheet}' tidak ditemukan.`);
    
    const originalData = sheet.getDataRange().getValues();

    if (originalData.length < 2) {
      return { headers: originalData.length > 0 ? originalData[0] : [], rows: [] };
    }
    
    const originalHeaders = originalData[0].map(h => String(h).trim());
    const dataRows = originalData.slice(1);

    const timestampIndex = originalHeaders.indexOf('Tanggal Unggah');
    const updateIndex = originalHeaders.indexOf('Update');
    const tglSKIndex = originalHeaders.indexOf('Tanggal SK');
    
    const sortIndex = (timestampIndex > -1) ? timestampIndex : 0;

    const parseDate = (value) => {
        if (!value) return new Date(0);
        if (value instanceof Date && !isNaN(value)) return value;
        const date = new Date(value);
        return isNaN(date) ? new Date(0) : date;
    };
    
    const indexedData = dataRows.map((row, index) => ({ row: row, originalIndex: index + 2 }));
    
    indexedData.sort((a, b) => {
      const dateB_update = (updateIndex > -1) ? parseDate(b.row[updateIndex]) : new Date(0);
      const dateA_update = (updateIndex > -1) ? parseDate(a.row[updateIndex]) : new Date(0);
      
      if (dateB_update.getTime() !== dateA_update.getTime()) {
        return dateB_update - dateA_update;
      }
      
      const dateB_timestamp = parseDate(b.row[sortIndex]);
      const dateA_timestamp = parseDate(a.row[sortIndex]);
      return dateB_timestamp - dateA_timestamp;
    });

    const formattedRows = indexedData.map(item => {
      const rowData = {};
      originalHeaders.forEach((header, i) => {
        let cell = item.row[i];
        if ((i === timestampIndex || i === updateIndex || i === tglSKIndex) && cell) {
          const dateObject = parseDate(cell);
          if (dateObject.getTime() !== 0) {
            const format = (i === tglSKIndex) ? "dd/MM/yyyy" : "dd/MM/yyyy HH:mm:ss";
            rowData[header] = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), format);
          } else {
            rowData[header] = '';
          }
        } else {
          rowData[header] = String(cell);
        }
      });
      return { rowIndex: item.originalIndex, data: rowData };
    });

    return { headers: originalHeaders, rows: formattedRows };

  } catch (e) {
    return handleError("getSKDataForManagement", e);
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

function getPaudSchoolLists() {
  const cacheKey = 'paud_school_lists';
  return getCachedData(cacheKey, function() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DATA_SEKOLAH_PAUD.id);
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DATA_SEKOLAH_PAUD.id); 
    
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

    const config = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;
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
    // PERBAIKAN: Menggunakan nama konfigurasi yang benar (RESPONSES bukan RESPONCES)
    const config = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const getValue = (key) => formData[key] || 0;

    // PERBAIKAN: Susun baris baru sesuai urutan kolom A s/d GH
    const newRow = [
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
      getValue('k4_agama_hindu_L'), getValue('k4_agama_hindu_P'), getValue('k4_agama_buddha_L'), getValue('k4_agama_buddha_P'), getValue('k4_agama_konghucu_L'), getValue('k4_agama_konghucu_P'),

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
    
    sheet.appendRow(newRow);

    return "Sukses! Laporan Bulan SD berhasil dikirim.";
  } catch (e) {
    return handleError('processLapbulFormSd', e);
  }
}

function getLapbulRiwayatData() {
  try {
    const paudSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
    const sdSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);

    const combinedHeaders = ["Tanggal Unggah", "Bulan", "Tahun", "Jenjang", "Nama Sekolah", "Status", "Rombel", "Dokumen"];
    let combinedData = [];

    const parseDate = (dateString) => {
        if (!dateString || typeof dateString !== 'string') return null;
        const parts = dateString.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s(\d{1,2}):(\d{2}):(\d{2})/);
        if (parts) return new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5], parts[6]);
        const isoDate = new Date(dateString);
        return isNaN(isoDate) ? null : isoDate;
    };

    const processSheetData = (sheet, jenjangDefault) => {
      if (!sheet) return;
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length < 2) return;

      const headers = data[0].map(h => h.trim());
      const rows = data.slice(1);

      // --- PERBAIKAN DI SINI ---
      const mapping = {
        'PAUD': { waktu: 'Tanggal Unggah', bulan: 'Bulan', tahun: 'Tahun', jenjang: 'Jenjang', nama: 'Nama Sekolah', status: 'Status', rombel: 'Jumlah Rombel', doc: 'Dokumen' },
        'SD': { waktu: 'Tanggal Unggah', bulan: 'Bulan', tahun: 'Tahun', jenjang: 'Jenjang', nama: 'Nama Sekolah', status: 'Status', rombel: 'Jumlah Rombel', doc: 'Dokumen' }
      };
      
      const currentMap = mapping[jenjangDefault];
      const colIndices = {};
      for (const key in currentMap) {
        const index = headers.indexOf(currentMap[key]);
        if (index === -1) throw new Error(`Header '${currentMap[key]}' tidak ditemukan di sheet '${jenjangDefault}'.`);
        colIndices[key] = index;
      }

      rows.forEach(row => {
        if (colIndices.waktu > -1 && row[colIndices.waktu]) {
          const rowData = [
            row[colIndices.waktu],
            row[colIndices.bulan],
            row[colIndices.tahun],
            row[colIndices.jenjang], // Dibaca langsung dari kolom 'Jenjang'
            row[colIndices.nama],
            row[colIndices.status],
            row[colIndices.rombel],
            row[colIndices.doc]
          ];
          rowData.push(parseDate(row[colIndices.waktu])); 
          combinedData.push(rowData);
        }
      });
    };

    processSheetData(paudSheet, 'PAUD');
    processSheetData(sdSheet, 'SD');

    combinedData.sort((a, b) => {
        const dateA = a[a.length - 1];
        const dateB = b[b.length - 1];
        if (!dateA) return 1; if (!dateB) return -1;
        return dateB - dateA;
    });
    
    const finalData = combinedData.map(row => row.slice(0, -1));

    return [combinedHeaders].concat(finalData);

  } catch(e) {
    Logger.log(`Error in getLapbulRiwayatData: ${e.message}\nStack: ${e.stack}`);
    return { error: `Terjadi error di server: ${e.message}` };
  }
}

function getLapbulStatusData() {
  try {
    const sheet = SpreadsheetApp.openById("1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU").getSheetByName("Status");
    if (!sheet) {
      throw new Error("Sheet 'Status' tidak ditemukan.");
    }
    const allData = sheet.getDataRange().getDisplayValues();
    if (allData.length < 2) {
      return { headers: [], rows: [], filters: {} };
    }

    const headers = allData[0];
    const dataRows = allData.slice(1);

    // Ekstrak data unik untuk filter
    const uniqueTahun = [...new Set(dataRows.map(row => row[1]))].filter(Boolean).sort().reverse(); // Kolom B
    const uniqueJenjang = [...new Set(dataRows.map(row => row[0]))].filter(Boolean).sort(); // Kolom A
    const uniqueStatus = [...new Set(dataRows.map(row => row[3]))].filter(Boolean).sort(); // Kolom D

    // Ambil header dari kolom C (indeks 2) sampai P (indeks 15)
    const displayHeaders = headers.slice(2, 16);
    
    // Proses baris data untuk dikirim ke client
    const displayRows = dataRows.map(row => {
      return {
        // Data ini digunakan untuk filtering di sisi client
        filters: {
          jenjang: row[0], // Kolom A
          tahun: row[1],   // Kolom B
          status: row[3]   // Kolom D
        },
        // Data ini yang akan ditampilkan di tabel (Kolom C s/d P)
        values: row.slice(2, 16)
      };
    });

    return {
      headers: displayHeaders,
      rows: displayRows,
      filters: {
        tahun: uniqueTahun,
        jenjang: uniqueJenjang,
        status: uniqueStatus
      }
    };
  } catch (e) {
    return handleError('getLapbulStatusData', e);
  }
}

function getLapbulKelolaData() {
  try {
    const paudSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName("Form Responses 1");
    const sdSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName("Input");

    const finalHeaders = ["Tanggal Unggah", "Bulan", "Tahun", "Jenjang", "Nama Sekolah", "Status", "Jumlah Rombel", "Dokumen", "Update"];
    let combinedData = [];

    const processSheetData = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      
      // --- PERBAIKAN UTAMA DI SINI: Gunakan getValues() ---
      const data = sheet.getDataRange().getValues(); // Mengambil nilai mentah, termasuk objek Date
      const headers = data[0].map(h => String(h).trim());
      
      const idx = {
        'PAUD': { ts: 0, bulan: 1, tahun: 2, jenjang: 6, nama: 7, status: 4, rombel: 5, doc: 36, update: 50 },
        'SD':   { ts: 0, bulan: 1, tahun: 2, jenjang: 190, nama: 4, status: 3, rombel: 6, doc: 7, update: 191 }
      };
      
      const currentIdx = idx[sourceName];
      const rows = data.slice(1);

      rows.forEach((row, index) => {
        const timestampCell = row[currentIdx.ts];
        if (!timestampCell) return;

        // Fungsi untuk memformat sel tanggal dengan aman
        const formatDate = (cell) => {
            if (cell instanceof Date) {
                return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
            }
            return cell; // Kembalikan nilai aslinya jika bukan objek Date
        };

        const timestamp = formatDate(timestampCell);
        const updateTimestamp = formatDate(row[currentIdx.update]);

        // Buat array data mentah
        const rowData = [
          timestamp,
          row[currentIdx.bulan],
          row[currentIdx.tahun],
          row[currentIdx.jenjang],
          row[currentIdx.nama],
          row[currentIdx.status],
          row[currentIdx.rombel],
          row[currentIdx.doc],
          updateTimestamp, // Gunakan nilai yang sudah diformat
          sourceName,
          index + 2
        ];
        combinedData.push(rowData);
      });
    };

    processSheetData(paudSheet, 'PAUD');
    processSheetData(sdSheet, 'SD');
    
    return { headers: finalHeaders, rows: combinedData };

  } catch (e) {
    Logger.log(`Error in getLapbulKelolaData: ${e.message}`);
    return { error: `Terjadi error di server: ${e.message}` };
  }
}

/**
 * Mengambil satu baris data Laporan Bulan berdasarkan sumber dan nomor baris.
 * @param {number} rowIndex Nomor baris di spreadsheet.
 * @param {string} source Sumber data ('PAUD' atau 'SD').
 * @returns {Object} Objek yang berisi data dari baris tersebut.
 */
function getLapbulDataByRow(rowIndex, source) {
  try {
    let sheet;
    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header.trim()] = values[i];
    });
    
    return rowData;
  } catch (e) {
    return handleError('getLapbulDataByRow', e);
  }
}


/**
 * Memperbarui satu baris data Laporan Bulan.
 * @param {Object} formData Objek data dari form edit di client.
 * @returns {string} Pesan sukses.
 */
function updateLapbulData(formData) {
  try {
    let sheet, FOLDER_ID;
    const source = formData.source;
    const rowIndex = formData.rowIndex;

    if (!source || !rowIndex) throw new Error("Informasi 'source' atau 'rowIndex' tidak ditemukan.");

    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
      // Dapatkan jenjang dari data yang dikirim jika ada, jika tidak, baca dari sheet
      const jenjang = formData.jenjang || sheet.getRange(rowIndex, headers.indexOf('Jenjang') + 1).getValue();
      FOLDER_ID = jenjang === 'KB' ? FOLDER_CONFIG.LAPBUL_KB : FOLDER_CONFIG.LAPBUL_TK;
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
      FOLDER_ID = FOLDER_CONFIG.LAPBUL_SD;
    } else {
      throw new Error("Sumber data tidak valid: " + source);
    }
    
    if (!formData.laporanBulan || !formData.tahun) {
        throw new Error("Informasi 'laporanBulan' atau 'tahun' kosong. Tidak dapat memproses file.");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
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
      const tahunFolder = getOrCreateFolder(mainFolder, formData.tahun);
      const bulanFolder = getOrCreateFolder(tahunFolder, formData.laporanBulan);
      
      const namaSekolah = formData.namaSekolah || existingValues[headers.indexOf('Nama Sekolah')];
      const newFileName = `${namaSekolah} - Lapbul ${formData.laporanBulan} ${formData.tahun}.pdf`;
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFileName);
      const newFile = bulanFolder.createFile(blob);
      formData[docHeaderName] = newFile.getUrl();
    }

    formData['Update'] = new Date();
    
    const newRowValues = headers.map((header, index) => {
      if (formData.hasOwnProperty(header)) {
        return formData[header];
      }
      return existingValues[index];
    });

    range.setValues([newRowValues]);

    // --- PERBAIKAN UTAMA DI SINI ---
    // 1. Cari posisi kolom "Update"
    const updateColIndex = headers.indexOf('Update');
    if (updateColIndex !== -1) {
      // 2. Set format sel di kolom tersebut agar selalu menampilkan tanggal dan jam
      sheet.getRange(rowIndex, updateColIndex + 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
    
    return "Data berhasil diperbarui.";
  } catch (e) {
    Logger.log(`Error in updateLapbulData: ${e.message}\nStack: ${e.stack}`);
    // Mengembalikan objek error agar bisa ditangkap di client
    return { error: `Terjadi error di server: ${e.message}` };
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

      // Ambil data dari B2 sampai baris terakhir
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

function getKeadaanPtkPaudData() {
  try {
    // ID Spreadsheet dari URL yang Anda berikan
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Keadaan PTK PAUD");
    
    if (!sheet) {
      throw new Error("Sheet 'Keadaan PTK PAUD' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();

  } catch (e) {
    return handleError('getKeadaanPtkPaudData', e);
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

/**
 * ===================================================================
 * ================= MODUL SK PEMBAGIAN TUGAS =====================
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

    // Mengubah setiap baris (yang tadinya array) menjadi objek
    const structuredRows = dataRows.map(row => {
      const rowObject = {};
      headers.forEach((header, index) => {
        rowObject[header] = row[index];
      });
      return rowObject;
    });
    
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

    // Fungsi bantu yang andal untuk mengubah berbagai format tanggal menjadi objek Date
    const parseDate = (value) => {
      if (value instanceof Date && !isNaN(value)) return value;
      if (!value) return new Date(0); 
      const date = new Date(value);
      return isNaN(date) ? new Date(0) : date;
    };

    // Simpan nomor baris asli SEBELUM diurutkan
    const indexedData = dataRows.map((row, index) => ({
      row: row,
      originalIndex: index + 2 // Baris fisik di spreadsheet (dimulai dari 2)
    }));

    // Logika pengurutan multi-level (sudah bagus, kita pertahankan)
    const updateIndex = originalHeaders.indexOf('Update');
    const timestampIndex = originalHeaders.indexOf('Tanggal Unggah');

    indexedData.sort((a, b) => {
      // Prioritas 1: Urutkan berdasarkan kolom "Update" (terbaru di atas)
      const dateB_update = parseDate(b.row[updateIndex]);
      const dateA_update = parseDate(a.row[updateIndex]);
      if (dateB_update.getTime() !== dateA_update.getTime()) {
        return dateB_update - dateA_update;
      }
      // Prioritas 2: Jika "Update" sama, urutkan berdasarkan "Tanggal Unggah"
      const dateB_timestamp = parseDate(b.row[timestampIndex]);
      const dateA_timestamp = parseDate(a.row[timestampIndex]);
      return dateB_timestamp - dateA_timestamp;
    });

    // Ubah setiap baris menjadi objek dan format tanggalnya
    const structuredRows = indexedData.map(item => {
      const rowObject = {
        _rowIndex: item.originalIndex, // WAJIB: untuk tombol Edit & Hapus
        _source: 'SK'                  // WAJIB: untuk membedakan dengan modul lain
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
    
    // Tentukan urutan header yang akan ditampilkan di tabel
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
    // ================== PERBAIKAN DI BARIS INI ==================
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES; // Nama variabel sudah benar
    // ==========================================================

    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const range = sheet.getRange(formData.rowIndex, 1, 1, headers.length);
    const existingRowValues = range.getDisplayValues()[0];
    const existingRowObject = {};
    headers.forEach((header, i) => { existingRowObject[header] = existingRowValues[i]; });

    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
    const tahunAjaranFolderName = existingRowObject['Tahun Ajaran'].replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);
    
    let fileUrl = existingRowObject['Dokumen'];
    const fileUrlIndex = headers.indexOf('Dokumen');

    const newSemesterFolderName = formData['Semester'];
    const newTargetFolder = getOrCreateFolder(tahunAjaranFolder, newSemesterFolderName);
    const newFilename = `${existingRowObject['Nama SD']} - ${tahunAjaranFolderName} - ${newSemesterFolderName} - ${formData['Kriteria SK']}.pdf`;

    if (formData.fileData && formData.fileData.data) {
      if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        try {
          const fileId = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
          if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch (e) {
          Logger.log(`Gagal menghapus file lama saat upload baru: ${e.message}`);
        }
      }
      
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
      const newFile = newTargetFolder.createFile(blob);
      fileUrl = newFile.getUrl();

    } else if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        const fileIdMatch = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
        if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const file = DriveApp.getFileById(fileId);
            const currentFileName = file.getName();
            
            const parents = file.getParents();
            let needsMove = true; 

            if (parents.hasNext()) {
                const currentParentFolder = parents.next();
                if (currentFileName === newFilename && currentParentFolder.getName() === newSemesterFolderName) {
                    needsMove = false;
                }
            }

            if (needsMove) {
                file.moveTo(newTargetFolder);
                file.setName(newFilename);
                fileUrl = file.getUrl();
                Logger.log(`File dipindahkan ke folder '${newSemesterFolderName}' dan diubah namanya menjadi '${newFilename}'`);
            }
        }
    }
    
    formData['Dokumen'] = fileUrl;
    formData['Update'] = new Date();
    const newRowValuesForSheet = headers.map(header => {
      return formData.hasOwnProperty(header) ? formData[header] : existingRowObject[header];
    });

    sheet.getRange(formData.rowIndex, 1, 1, headers.length).setValues([newRowValuesForSheet]);
    
    const tanggalSKIndex = headers.indexOf('Tanggal SK');
    if (tanggalSKIndex !== -1) {
      sheet.getRange(formData.rowIndex, tanggalSKIndex + 1).setNumberFormat("dd-MM-yyyy");
    }
    
    SpreadsheetApp.flush(); 
    
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
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
 * [REFACTOR - FINAL] Mengambil data riwayat pengiriman SK.
 * Mengembalikan data dalam format objek { headers, rows } yang sudah diurutkan.
 */
function getSKRiwayatData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.id)
                                .getSheetByName(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.sheet);

    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Tanggal SK", "Kriteria SK", "Dokumen", "Tanggal Unggah"], rows: [] };
    }
    
    const allData = sheet.getDataRange().getValues();
    const originalHeaders = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    const parseDate = (value) => {
        if (!value) return new Date(0);
        if (value instanceof Date && !isNaN(value)) return value;
        const date = new Date(value);
        return isNaN(date) ? new Date(0) : date;
    };

    let structuredRows = dataRows.map(row => {
      const rowObject = {};
      originalHeaders.forEach((header, index) => {
        rowObject[header] = row[index];
      });
      return rowObject;
    });

    structuredRows.sort((a, b) => {
      const dateB = parseDate(b['Tanggal Unggah']);
      const dateA = parseDate(a['Tanggal Unggah']);
      return dateB - dateA;
    });
    
    const formattedRows = structuredRows.map(row => {
        if (row['Tanggal Unggah'] instanceof Date) {
          row['Tanggal Unggah'] = Utilities.formatDate(row['Tanggal Unggah'], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        }
        if (row['Tanggal SK'] instanceof Date) {
          row['Tanggal SK'] = Utilities.formatDate(row['Tanggal SK'], Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        return row;
    });
    
    const desiredHeaders = ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Tanggal SK", "Kriteria SK", "Dokumen", "Tanggal Unggah"];

    return {
      headers: desiredHeaders,
      rows: formattedRows
    };

  } catch (e) {
    return handleError('getSKRiwayatData', e);
  }
}

/**
 * [REFACTOR - FINAL] Mengambil data status pengiriman SK.
 * Mengembalikan data dalam format objek { headers, rows } agar sesuai dengan fungsi generik.
 */
function getSKStatusData() {
  try {
    // Memanggil fungsi utilitas yang ada di Code.gs
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
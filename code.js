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
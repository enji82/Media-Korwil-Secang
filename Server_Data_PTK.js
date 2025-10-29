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

/**
 * Mengambil data jumlah PTK bulanan untuk PAUD dari spreadsheet.
 */
function getJumlahPtkPaudBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Data PAUD Bulanan");
    if (!sheet) {
      throw new Error("Sheet 'Data PAUD Bulanan' tidak ditemukan.");
    }
    
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    // [PERBAIKAN] Menambahkan Kolom D (Bulan) dan E (Tahun) untuk filter
    // Indeks kolom yang diinginkan (A=0, B=1, C=2, D=3, E=4, F=5, dst.)
    const colIndices = [0, 1, 2, 3, 4, 5, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 35];
    // Filter header sesuai kolom yang diinginkan
    const finalHeaders = colIndices.map(index => headers[index]);

    // Filter setiap baris untuk hanya menyertakan data dari kolom yang diinginkan
    const finalData = dataRows.map(row => {
      return colIndices.map(index => row[index]);
    });

    return [finalHeaders].concat(finalData);

  } catch (e) {
    return handleError('getJumlahPtkPaudBulananData', e);
  }
}

/**
 * Mengambil data daftar PTK PAUD dari spreadsheet.
 */
function getDaftarPtkPaudData() {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) {
      throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");
    }
    
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    let dataRows = allData.slice(1);

    // [PERBAIKAN] Mengurutkan baris data berdasarkan abjad pada kolom "Nama" (indeks 2)
    dataRows.sort((a, b) => {
      // localeCompare digunakan untuk pengurutan abjad yang benar
      return a[2].localeCompare(b[2]);
    });

    const colIndices = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    const finalHeaders = colIndices.map(index => headers[index]);
    
    const finalData = dataRows.map(row => {
      return colIndices.map(index => row[index]);
    });

    return [finalHeaders].concat(finalData);

  } catch (e) {
    return handleError('getDaftarPtkPaudData', e);
  }
}

function getKelolaPtkPaudData() {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const allData = sheet.getDataRange().getValues();
    if (allData.length < 2) {
      return { headers: [], rows: [] };
    }

    const headers = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    // [PERBAIKAN] Langkah 1: Simpan nomor baris asli SEBELUM diurutkan
    const indexedData = dataRows.map((row, index) => ({
      originalRowIndex: index + 2, // Baris fisik di spreadsheet (mulai dari 2)
      rowData: row
    }));

    // Fungsi bantu untuk mengubah tanggal menjadi objek yang bisa dibandingkan
    const parseDate = (value) => {
      if (value instanceof Date && !isNaN(value)) {
        return value.getTime();
      }
      return 0; 
    };

    // Langkah 2: Lakukan pengurutan 2 tingkat pada data yang sudah diindeks
    const updateIndex = headers.indexOf('Update');
    const dateInputIndex = headers.indexOf('Tanggal Input');
    
    indexedData.sort((a, b) => {
      // Urutan 1: Berdasarkan kolom "Update" (terbaru di atas)
      const updateA = parseDate(a.rowData[updateIndex]);
      const updateB = parseDate(b.rowData[updateIndex]);
      if (updateB !== updateA) {
        return updateB - updateA;
      }

      // Urutan 2: Jika "Update" sama, urutkan berdasarkan "Tanggal Input" (terbaru di atas)
      const dateInputA = parseDate(a.rowData[dateInputIndex]);
      const dateInputB = parseDate(b.rowData[dateInputIndex]);
      return dateInputB - dateInputA;
    });

    // Langkah 3: Buat data final untuk dikirim ke frontend
    const finalData = indexedData.map(item => {
      const rowDataObject = {};
      headers.forEach((header, i) => {
        let cell = item.rowData[i];
        if (cell instanceof Date) {
          rowDataObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        } else {
          rowDataObject[header] = cell || "";
        }
      });
      
      // [KUNCI PERBAIKAN] Kirim nomor baris asli yang sudah disimpan sebelumnya
      return { rowIndex: item.originalRowIndex, data: rowDataObject }; 
    });

    return { headers, rows: finalData };
  } catch (e) {
    Logger.log(`Error getKelolaPtkPaudData: ${e.message}`);
    return { error: e.message, headers: [], rows: [] };
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
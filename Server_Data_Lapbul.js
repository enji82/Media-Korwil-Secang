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
    const config = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const getValue = (key) => formData[key] || 0;

    // Susun baris baru sesuai urutan kolom A s/d GH
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
    sheet.appendRow(newRow);

    return "Sukses! Laporan Bulan SD berhasil dikirim.";
  } catch (e) {
    return handleError('processLapbulFormSd', e);
  }
}

function getLapbulRiwayatData() {
  try {
    const sheet = SpreadsheetApp.openById("1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU").getSheetByName("Riwayat");

    if (!sheet) {
      throw new Error("Sheet 'Riwayat' di spreadsheet gabungan tidak ditemukan.");
    }

    const allData = sheet.getDataRange().getDisplayValues();

    // KUNCI PERUBAHAN 1: Tentukan urutan kolom final yang Anda inginkan di sini.
    // "Jenjang" tetap ada di akhir agar logika penyembunyian kolom di JavaScript tetap berfungsi.
    const desiredHeaders = ["Nama Sekolah", "Status", "Bulan", "Tahun", "Rombel", "Dokumen", "Tanggal Unggah", "Jenjang"];

    if (allData.length < 2) {
      return [desiredHeaders]; // Jika data kosong, kirim header dengan urutan yang benar.
    }

    const sourceHeaders = allData[0].map(h => h.trim());
    const dataRows = allData.slice(1);

    // KUNCI PERBAIKAN 2: Buat peta untuk menemukan posisi kolom asli di spreadsheet.
    const headerMap = {};
    sourceHeaders.forEach((header, index) => {
      headerMap[header] = index;
    });

    // KUNCI PERBAIKAN 3: Susun ulang setiap baris data sesuai urutan 'desiredHeaders'.
    const reorderedDataRows = dataRows.map(row => {
      const newRow = [];
      desiredHeaders.forEach(header => {
        const sourceIndex = headerMap[header];
        // Jika kolom ditemukan di sumber, ambil datanya. Jika tidak, beri placeholder '-'.
        newRow.push(sourceIndex !== undefined ? row[sourceIndex] : '-');
      });
      return newRow;
    });

    // Kirim header baru dan data yang sudah disusun ulang.
    // Proses pengurutan berdasarkan tanggal akan tetap dilakukan di sisi browser seperti sebelumnya.
    return [desiredHeaders].concat(reorderedDataRows);

  } catch(e) {
    Logger.log(`Error in getLapbulRiwayatData: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Terjadi error di server: ${e.message}`);
  }
}

function getLapbulStatusData() {
  try {
    // Membuka spreadsheet dan sheet "Status" yang benar.
    const sheet = SpreadsheetApp.openById("1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU").getSheetByName("Status");
    
    if (!sheet) {
      throw new Error("Sheet 'Status' tidak ditemukan.");
    }
    
    // Langsung ambil dan kirim semua data yang terlihat (getDisplayValues).
    // Proses pemilihan kolom dan filter akan dilakukan di javascript.html
    return sheet.getDataRange().getDisplayValues();

  } catch (e) {
    Logger.log(`Error in getLapbulStatusData: ${e.message}`);
    throw new Error(`Terjadi error di server: ${e.message}`);
  }
}

function getLapbulKelolaData() {
  try {
    const paudSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName("Form Responses 1");
    const sdSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName("Input");
    const finalHeaders = ["Nama Sekolah", "Status", "Jenjang", "Jumlah Rombel", "Bulan", "Tahun", "Dokumen", "Tanggal Unggah", "Update"];
    let combinedData = [];

    const parseDateForSort = (dateStr) => {
        if (!dateStr || !(typeof dateStr === 'string' || dateStr instanceof Date)) return new Date(0);
        if (dateStr instanceof Date) return dateStr;
        const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
        if (parts) {
            return new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5], parts[6]);
        }
        const dateOnlyParts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (dateOnlyParts) {
            return new Date(dateOnlyParts[3], dateOnlyParts[2] - 1, dateOnlyParts[1]);
        }
        return new Date(0);
    };

    const processSheetData = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());
      const rows = data.slice(1);

      const formatDate = (cell) => {
          if (cell instanceof Date) {
              return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          }
          return cell;
      };

      rows.forEach((row, index) => {
        const timestampCell = row[headers.indexOf("Tanggal Unggah")];
        if (!timestampCell) return;

        const rowData = {};
        headers.forEach((header, i) => {
            rowData[header] = (header === "Tanggal Unggah" || header === "Update") ? formatDate(row[i]) : row[i];
        });

        if (sourceName === 'SD') {
          rowData['Jenjang'] = 'SD';
        }
        
        combinedData.push({
          rowIndex: index + 2,
          source: sourceName,
          data: rowData
        });
      });
    };

    processSheetData(paudSheet, 'PAUD');
    processSheetData(sdSheet, 'SD');
    
    // KUNCI PERBAIKAN: Logika pengurutan diubah agar HANYA menggunakan "Tanggal Unggah"
    combinedData.sort((a, b) => {
        const dateB = parseDateForSort(b.data['Tanggal Unggah']);
        const dateA = parseDateForSort(a.data['Tanggal Unggah']);
        return dateB - dateA; // Mengurutkan dari yang terbaru ke terlama
    });

    return { headers: finalHeaders, rows: combinedData };

  } catch (e) {
    Logger.log(`Error in getLapbulKelolaData: ${e.message}\nStack: ${e.stack}`);
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

    // KUNCI PERBAIKAN: Blok ini dipindahkan ke atas agar variabel 'sheet'
    // terdefinisi sebelum digunakan.
    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
      const jenjang = formData.jenjang || sheet.getRange(rowIndex, headers.indexOf('Jenjang') + 1).getValue();
      FOLDER_ID = jenjang === 'KB' ? FOLDER_CONFIG.LAPBUL_KB : FOLDER_CONFIG.LAPBUL_TK;
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
      FOLDER_ID = FOLDER_CONFIG.LAPBUL_SD;
    } else {
      throw new Error("Sumber data tidak valid: " + source);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());

    if (!formData.laporanBulan || !formData.tahun) {
        throw new Error("Informasi 'laporanBulan' atau 'tahun' kosong. Tidak dapat memproses file.");
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

function getLapbulArsipFolderIds() {
  try {
    return {
      'KB': FOLDER_CONFIG.LAPBUL_KB,
      'TK': FOLDER_CONFIG.LAPBUL_TK,
      'SD': FOLDER_CONFIG.LAPBUL_SD
    };
  } catch (e) {
    return handleError('getLapbulArsipFolderIds', e);
  }
}
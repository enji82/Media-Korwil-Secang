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
 * [REFACTOR - FINAL] Mengambil data riwayat pengiriman Laporan Bulan.
 * Mengembalikan data dalam format objek yang sudah diurutkan dan siap pakai.
 */
function getLapbulRiwayatData() {
  try {
    const desiredHeaders = ["Nama Sekolah", "Status", "Jenjang", "Bulan", "Tahun", "Rombel", "Dokumen", "Tanggal Unggah"];
    let combinedData = [];

    const processSheetData = (sheet, mapping) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());
      const dataRows = data.slice(1);

      dataRows.forEach(row => {
        if (!row[0]) return;

        const rowObject = {
          "Nama Sekolah":   row[headers.indexOf(mapping.namaSekolah)],
          "Status":         row[headers.indexOf(mapping.status)],
          "Jenjang":        mapping.jenjang === 'static' ? 'SD' : row[headers.indexOf(mapping.jenjang)],
          "Bulan":          row[headers.indexOf(mapping.bulan)],
          "Tahun":          row[headers.indexOf(mapping.tahun)],
          "Rombel":         row[headers.indexOf(mapping.rombel)],
          "Dokumen":        row[headers.indexOf(mapping.dokumen)],
          "Tanggal Unggah": row[headers.indexOf(mapping.tanggalUnggah)]
        };
        combinedData.push(rowObject);
      });
    };
    
    // 1. Proses Data SD (Tidak ada perubahan)
    const sdConfig = SPREADSHEET_CONFIG.LAPBUL_SD;
    const sdSheet = SpreadsheetApp.openById(sdConfig.id).getSheetByName(sdConfig.sheet);
    const sdMapping = { 
        namaSekolah: 'Nama Sekolah', jenjang: 'static', status: 'Status', bulan: 'Bulan', 
        tahun: 'Tahun', rombel: 'Rombel', dokumen: 'Dokumen', tanggalUnggah: 'Tanggal Unggah' 
    };
    processSheetData(sdSheet, sdMapping);

    // 2. Proses Data PAUD
    const paudConfig = SPREADSHEET_CONFIG.LAPBUL_PAUD;
    const paudSheet = SpreadsheetApp.openById(paudConfig.id).getSheetByName(paudConfig.sheet);
    
    // ================== AWAL PERBAIKAN DI SINI ==================
    const paudMapping = { 
        namaSekolah: 'Nama Sekolah',
        jenjang: 'Jenjang', 
        status: 'Status', // Diperbaiki dari 'Status Sekolah' menjadi 'Status'
        bulan: 'Bulan', 
        tahun: 'Tahun', 
        rombel: 'Rombel', // Diperbaiki dari 'Jumlah Rombel' menjadi 'Rombel'
        dokumen: 'Dokumen', 
        tanggalUnggah: 'Tanggal Unggah' 
    };
    // =================== AKHIR PERBAIKAN DI SINI ===================
    processSheetData(paudSheet, paudMapping);
    
    // 3. Urutkan semua data (tidak ada perubahan)
    const parseDate = (value) => {
        if (value instanceof Date && !isNaN(value)) return value;
        return new Date(0);
    };
    combinedData.sort((a, b) => {
      const dateB = parseDate(b["Tanggal Unggah"]);
      const dateA = parseDate(a["Tanggal Unggah"]);
      return dateB - dateA;
    });

    // 4. Format tanggal (tidak ada perubahan)
    const formattedData = combinedData.map(row => {
      if (row["Tanggal Unggah"] instanceof Date) {
        row["Tanggal Unggah"] = Utilities.formatDate(row["Tanggal Unggah"], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      }
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
    const sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_GABUNGAN.id).getSheetByName("Status");
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }
    
    const allData = sheet.getDataRange().getDisplayValues();
    const sourceHeaders = allData[0];
    const dataRows = allData.slice(1);

    // Tentukan kolom mana yang ingin ditampilkan dari spreadsheet sumber (indeks dimulai dari 0)
    // C=2, E=4, F=5, ..., N=13, O=14, P=15
    const displayIndices = [2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]; 
    const finalHeaders = displayIndices.map(index => sourceHeaders[index]);

    // Ubah setiap baris menjadi objek dan tambahkan data filter tersembunyi
    const structuredRows = dataRows.map(row => {
      const rowObject = {
        // Data tersembunyi untuk keperluan filter di sisi klien
        _filterJenjang: row[0], // Kolom A (Jenjang)
        _filterTahun: row[1],   // Kolom B (Tahun)
        _filterStatus: row[3]   // Kolom D (Status Sekolah)
      };
      
      // Isi data utama untuk ditampilkan di tabel
      finalHeaders.forEach((header, index) => {
        const displayIndex = displayIndices[index];
        rowObject[header] = row[displayIndex];
      });
      
      return rowObject;
    });

    return {
      headers: finalHeaders,
      rows: structuredRows
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
    // ================== AWAL PERBAIKAN DI SINI ==================
    // Menggunakan nama konfigurasi yang benar dari Config.gs, sesuai file code.js
    const paudConfig = SPREADSHEET_CONFIG.LAPBUL_PAUD;
    const sdConfig = SPREADSHEET_CONFIG.LAPBUL_SD;
    // =================== AKHIR PERBAIKAN DI SINI ===================

    const paudSheet = SpreadsheetApp.openById(paudConfig.id).getSheetByName(paudConfig.sheet);
    const sdSheet = SpreadsheetApp.openById(sdConfig.id).getSheetByName(sdConfig.sheet);
    
    let combinedData = [];

    // Fungsi bantu (tidak berubah)
    const parseDate = (value) => {
        if (!value) return new Date(0); 
        if (value instanceof Date && !isNaN(value)) return value;
        if (typeof value === 'string') {
          const parts = value.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
          if (parts) { return new Date(parts[3], parseInt(parts[2], 10) - 1, parts[1], parts[4], parts[5], parts[6]); }
        }
        return new Date(0);
    };

    const processSheetData = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
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
          rowObject[header] = row[i];
        });
        if (sourceName === 'SD') {
          // Set Jenjang secara manual untuk sumber data SD
          rowObject['Jenjang'] = 'SD';
        }
        combinedData.push(rowObject);
      });
    };

    processSheetData(paudSheet, 'PAUD');
    processSheetData(sdSheet, 'SD');
    
    // Blok pengurutan (tidak berubah)
    combinedData.sort((a, b) => {
        const dateB_update = parseDate(b['Update']);
        const dateA_update = parseDate(a['Update']);
        if (dateB_update.getTime() !== dateA_update.getTime()) {
            return dateB_update - dateA_update;
        }
        const dateB_timestamp = parseDate(b['Tanggal Unggah']);
        const dateA_timestamp = parseDate(a['Tanggal Unggah']);
        return dateB_timestamp - dateA_timestamp;
    });
    
    // Format tanggal menjadi string (tidak berubah)
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
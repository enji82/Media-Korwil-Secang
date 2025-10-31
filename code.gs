/**
 * ===================================================================
 * ======================= 1. KONFIGURASI PUSAT ======================
 * ===================================================================
 */
const SPREADSHEET_CONFIG = {
  // --- Modul SK Pembagian Tugas ---
  SK_BAGI_TUGAS: { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "SK Tabel Kirim" },
  SK_FORM_RESPONSES: { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "Form Responses 1" },

  // --- Modul Laporan Bulanan & Data Murid ---
  LAPBUL_FORM_RESPONSES_PAUD: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Form Responses 1" },
  LAPBUL_FORM_RESPONSES_SD: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Input" },
  LAPBUL_GABUNGAN: { id: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU" },

  // --- Modul Data PTK ---
  PTK_PAUD_KEADAAN: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Keadaan PTK PAUD" },
  PTK_PAUD_JUMLAH_BULANAN: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Data PAUD Bulanan" },
  PTK_PAUD_DB: { id: "1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk", sheet: "PTK PAUD" },
  PTK_SD_KEADAAN: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Keadaan PTK SD" },
  PTK_SD_JUMLAH_BULANAN: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "PTK Bulanan SD"},
  PTK_SD_KEBUTUHAN: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Kebutuhan Guru"},
  PTK_SD_DB: { id: "1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0" },

  // --- Modul Data Murid ---
  MURID_PAUD_KELAS: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Murid Kelas" },
  MURID_PAUD_JK: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Murid JK" },
  MURID_PAUD_BULANAN: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Murid Bulanan" },
  MURID_SD_KELAS: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "SD Tabel Kelas" },
  MURID_SD_ROMBEL: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "SD Tabel Rombel" },
  MURID_SD_JK: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "SD Tabel JK" },
  MURID_SD_AGAMA: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "SD Tabel Agama" },
  MURID_SD_BULANAN: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "SD Tabel Bulanan" },

  // --- Data Pendukung & Dropdown ---
  DATA_SEKOLAH: { id: "1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs" },   
  DROPDOWN_DATA: { id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA" },
  FORM_OPTIONS_DB: { id: "1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU" },

  // --- Data SIABA ---
  SIABA_REKAP: { id: "1x3b-yzZbiqP2XfJNRC3XTbMmRTHLd8eEdUqAlKY3v9U", sheet: "Rekap Script" },
  SIABA_TIDAK_PRESENSI: { id: "1mjXz5l_cqBiiR3x9qJ7BU4yQ3f0ghERT9ph8CC608Zc", sheet: "Rekap Script" },
};

const FOLDER_CONFIG = {
  MAIN_SK: "1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs",
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
  return { error: error.message }; // Return object error ke client
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function getDataFromSheet(configKey) {
  const config = SPREADSHEET_CONFIG[configKey];
  if (!config) throw new Error(`Konfigurasi untuk '${configKey}' tidak ditemukan.`);
  const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
  if (!sheet) throw new Error(`Sheet '${config.sheet}' di spreadsheet '${config.id}' tidak ditemukan.`);
  return sheet.getDataRange().getDisplayValues();
}

function getCachedData(key, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached != null) {
    return JSON.parse(cached);
  }
  const freshData = fetchFunction();
  cache.put(key, JSON.stringify(freshData), 21600); // Cache for 6 hours
  return freshData;
}


/**
 * ===================================================================
 * =================== MODUL GOOGLE DRIVE (ARSIP) ====================
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


/**
 * ===================================================================
 * ==================== MODUL DATA LAPORAN BULAN =====================
 * ===================================================================
 */

function getPaudSchoolLists() {
  const cacheKey = 'paud_school_lists_final_v4'; // Kunci cache baru
  return getCachedData(cacheKey, function() {
    
    // 1. Menggunakan kunci DROPDOWN_DATA (sesuai ID: 1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA)
    const config = SPREADSHEET_CONFIG.DROPDOWN_DATA; 
    
    if (!config || !config.id) {
        throw new Error("Konfigurasi ID Spreadsheet (DROPDOWN_DATA) tidak ditemukan.");
    }

    const ss = SpreadsheetApp.openById(config.id);
    if (!ss) {
        throw new Error("Gagal membuka Spreadsheet. Periksa ID atau izin akses.");
    }
    
    // 2. Gunakan nama Sheet yang benar: Form PAUD
    const sheet = ss.getSheetByName('Form PAUD');
    if (!sheet) {
        throw new Error("Sheet 'Form PAUD' tidak ditemukan di Spreadsheet Dropdown Data.");
    }
    
    if (sheet.getLastRow() < 2) {
        return { "KB": [], "TK": [] };
    }

    // 3. Ambil data Jenjang (Kolom A) dan Nama Lembaga (Kolom B)
    // Ambil data mulai dari baris 2, kolom 1 (A) sepanjang 2 kolom (A & B)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
    
    const lists = { "KB": [], "TK": [] };

    data.forEach(row => {
        const jenjang = String(row[0]).trim();
        const namaLembaga = String(row[1]).trim();
        
        if (jenjang && namaLembaga) {
            // Hanya tambahkan ke array yang sudah didefinisikan (KB atau TK)
            if (lists.hasOwnProperty(jenjang) && !lists[jenjang].includes(namaLembaga)) {
                lists[jenjang].push(namaLembaga);
            }
        }
    });
    
    // Urutkan Nama Lembaga
    lists["KB"].sort();
    lists["TK"].sort();

    return lists;
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
    const desiredHeaders = ["Nama Sekolah", "Status", "Bulan", "Tahun", "Rombel", "Dokumen", "Tanggal Unggah", "Jenjang"];

    if (allData.length < 2) {
      return { headers: desiredHeaders, rows: [] };
    }

    const sourceHeaders = allData[0].map(h => h.trim());
    const dataRows = allData.slice(1);
    const headerMap = {};
    sourceHeaders.forEach((header, index) => {
      headerMap[header] = index;
    });

    const reorderedDataRows = dataRows.map(row => {
      const newRow = {};
      desiredHeaders.forEach(header => {
        const sourceIndex = headerMap[header];
        newRow[header] = (sourceIndex !== undefined ? row[sourceIndex] : '-');
      });
      return newRow;
    });
    
    // Urutkan berdasarkan Tanggal Unggah (asumsi formatnya bisa di-parse)
    reorderedDataRows.sort((a, b) => {
        const dateA = new Date(a['Tanggal Unggah'].split(' ')[0].split('/').reverse().join('-'));
        const dateB = new Date(b['Tanggal Unggah'].split(' ')[0].split('/').reverse().join('-'));
        return dateB - dateA;
    });

    return { headers: desiredHeaders, rows: reorderedDataRows };

  } catch(e) {
    return handleError('getLapbulRiwayatData', e);
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
        return { headers: [], rows: [], filterConfigs: [] };
    }
    
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const jenjangIndex = headers.indexOf('Jenjang');
    const tahunIndex = headers.indexOf('Tahun');
    
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const currentMonthName = monthNames[new Date().getMonth()];

    const processedRows = dataRows.map(row => {
        const rowObject = {};
        headers.forEach((h, i) => rowObject[h] = row[i]);
        
        rowObject['_filterJenjang'] = row[jenjangIndex];
        rowObject['_filterTahun'] = row[tahunIndex];
        
        const submittedMonths = headers.slice(2).filter((h, i) => row[i + 2] === 'T').map(h => h.split(' ')[0]);
        if (submittedMonths.includes(currentMonthName)) {
            rowObject['_filterStatus'] = 'Sudah Mengirim';
        } else {
            rowObject['_filterStatus'] = 'Belum Mengirim';
        }
        return rowObject;
    });

    return {
        headers: headers,
        rows: processedRows,
        filterConfigs: [
            { id: 'filterTahun', dataColumn: '_filterTahun', sortReverse: true },
            { id: 'filterJenjang', dataColumn: '_filterJenjang' },
            { id: 'filterStatus', dataColumn: '_filterStatus' }
        ]
    };

  } catch (e) {
    return handleError('getLapbulStatusData', e);
  }
}

function getLapbulKelolaData() {
  try {
    const paudSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName("Form Responses 1");
    const sdSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName("Input");
    const finalHeaders = ["Nama Sekolah", "Status", "Jenjang", "Jumlah Rombel", "Bulan", "Tahun", "Dokumen", "Tanggal Unggah", "Update", "Aksi"];
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
        const timestampCell = row[headers.indexOf("Timestamp")] || row[headers.indexOf("Tanggal Unggah")];
        if (!timestampCell) return;

        const rowData = {
          _rowIndex: index + 2,
          _source: sourceName
        };

        headers.forEach((header, i) => {
            const cleanHeader = header === "Timestamp" ? "Tanggal Unggah" : header;
            rowData[cleanHeader] = (cleanHeader === "Tanggal Unggah" || cleanHeader === "Update") ? formatDate(row[i]) : row[i];
        });

        if (sourceName === 'SD') {
          rowData['Jenjang'] = 'SD';
          rowData['Jumlah Rombel'] = rowData['Rombel'];
        }
        combinedData.push(rowData);
      });
    };

    processSheetData(paudSheet, 'PAUD');
    processSheetData(sdSheet, 'SD');
    
    combinedData.sort((a, b) => {
        const dateB = parseDateForSort(b['Tanggal Unggah']);
        const dateA = parseDateForSort(a['Tanggal Unggah']);
        return dateB - dateA;
    });

    return { headers: finalHeaders, rows: combinedData };

  } catch (e) {
    return handleError('getLapbulKelolaData', e);
  }
}

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

function updateLapbulData(formData) {
  try {
    let sheet, FOLDER_ID;
    const source = formData.source;
    const rowIndex = formData.rowIndex;
    if (!source || !rowIndex) throw new Error("Informasi 'source' atau 'rowIndex' tidak ditemukan.");

    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
      const jenjang = formData.Jenjang || sheet.getRange(rowIndex, headers.indexOf('Jenjang') + 1).getValue();
      FOLDER_ID = jenjang === 'KB' ? FOLDER_CONFIG.LAPBUL_KB : FOLDER_CONFIG.LAPBUL_TK;
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
      FOLDER_ID = FOLDER_CONFIG.LAPBUL_SD;
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
    return { error: `Terjadi error di server: ${e.message}` };
  }
}

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
  const cacheKey = 'lapbul_info_v2'; // Ubah kunci cache agar selalu memuat data baru
  return getCachedData(cacheKey, function() {
    try {
      // Menggunakan SPREADSHEET_CONFIG.DROPDOWN_DATA (ID: 1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA)
      const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
      const sheet = ss.getSheetByName('Informasi');
      
      if (!sheet || sheet.getLastRow() < 2) {
        return []; // Kembalikan array kosong jika sheet tidak valid atau hanya header
      }

      const lastRow = sheet.getLastRow();
      // Ambil data dari A2 sampai baris terakhir
      const range = sheet.getRange('A2:A' + lastRow);
      const values = range.getValues()
                          .flat()
                          .filter(item => String(item).trim() !== ''); // Filter baris kosong
      
      // Jika Anda ingin mengambil data dari kolom B, gunakan range 'B2:B' + lastRow. 
      // Saya asumsikan Anda ingin kolom A (Informasi Umum) di sini.

      return values;
    } catch (e) {
      Logger.log(`Error in getLapbulInfo fetch: ${e.message}`);
      return []; // Kembalikan array kosong untuk menghindari error di client
    }
  });
}

function getUnduhFormatInfo() {
  const cacheKey = 'unduh_format_info_v1';
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
      const sheet = ss.getSheetByName('Informasi');
      if (!sheet || sheet.getLastRow() < 2) return [];
      const range = sheet.getRange('B2:B' + sheet.getLastRow());
      return range.getDisplayValues().flat().filter(item => String(item).trim() !== '');
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

/**
 * ===================================================================
 * ======================== MODUL: DATA MURID ========================
 * ===================================================================
 */

function getMuridPaudKelasUsiaData() {
  try {
    return getDataFromSheet('MURID_PAUD_KELAS');
  } catch (e) {
    return handleError('getMuridPaudKelasUsiaData', e);
  }
}

function getMuridPaudJenisKelaminData() {
  try {
    return getDataFromSheet('MURID_PAUD_JK');
  } catch (e) {
    return handleError('getMuridPaudJenisKelaminData', e);
  }
}

function getMuridPaudJumlahBulananData() {
  try {
    return getDataFromSheet('MURID_PAUD_BULANAN');
  } catch (e) {
    return handleError('getMuridPaudJumlahBulananData', e);
  }
}

function getMuridSdKelasData() {
  try {
    return getDataFromSheet('MURID_SD_KELAS');
  } catch (e) {
    return handleError('getMuridSdKelasData', e);
  }
}

function getMuridSdRombelData() {
  try {
    return getDataFromSheet('MURID_SD_ROMBEL');
  } catch (e) {
    return handleError('getMuridSdRombelData', e);
  }
}

function getMuridSdJenisKelaminData() {
  try {
    return getDataFromSheet('MURID_SD_JK');
  } catch (e) {
    return handleError('getMuridSdJenisKelaminData', e);
  }
}

function getMuridSdAgamaData() {
  try {
    return getDataFromSheet('MURID_SD_AGAMA');
  } catch (e) {
    return handleError('getMuridSdAgamaData', e);
  }
}

function getMuridSdJumlahBulananData() {
  try {
    const config = SPREADSHEET_CONFIG.MURID_SD_BULANAN;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan");
    return sheet.getDataRange().getValues(); // Menggunakan getValues() untuk angka
  } catch (e) {
    return handleError('getMuridSdJumlahBulananData', e);
  }
}

/**
 * ===================================================================
 * ========================= MODUL: DATA PTK =========================
 * ===================================================================
 */
 
function getPtkKeadaanPaudData() {
    try {
        const data = getDataFromSheet('PTK_PAUD_KEADAAN');
        if (!data || data.length < 2) {
            return { headers: [], rows: [], filterConfigs: [] };
        }
        
        const headers = data[0];
        const dataRows = data.slice(1);
        const jenjangIndex = headers.indexOf('Jenjang');

        const processedRows = dataRows.map(row => {
            const rowObject = {};
            headers.forEach((h, i) => rowObject[h] = row[i]);
            rowObject['_filterJenjang'] = row[jenjangIndex];
            return rowObject;
        });

        return { 
            headers: headers, 
            rows: processedRows,
            filterConfigs: [{ id: 'filterJenjang', dataColumn: '_filterJenjang' }]
        };
    } catch (e) {
        return handleError('getPtkKeadaanPaudData', e);
    }
}
 
function getPtkKeadaanSdData() {
  try {
    return getDataFromSheet('PTK_SD_KEADAAN');
  } catch (e) {
    return handleError('getKeadaanPtkSdData', e);
  }
}

function getPtkJumlahBulananPaudData() {
  try {
    const allData = getDataFromSheet('PTK_PAUD_JUMLAH_BULANAN');
    if (!allData || allData.length < 2) {
      return { headers: [], rows: [], filterConfigs: [] };
    }
    
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const jenjangIndex = headers.indexOf('Jenjang');
    const tahunIndex = headers.indexOf('Tahun');
    const bulanIndex = headers.indexOf('Bulan');
    
    const processedRows = dataRows.map(row => {
        const rowObject = {};
        headers.forEach((h, i) => rowObject[h] = row[i]);
        rowObject['_filterJenjang'] = row[jenjangIndex];
        rowObject['_filterTahun'] = row[tahunIndex];
        rowObject['_filterBulan'] = row[bulanIndex];
        return rowObject;
    });

    return {
        headers: headers,
        rows: processedRows,
        filterConfigs: [
            { id: 'filterTahun', dataColumn: '_filterTahun', sortReverse: true },
            { id: 'filterBulan', dataColumn: '_filterBulan', specialSort: 'bulan' },
            { id: 'filterJenjang', dataColumn: '_filterJenjang' },
            { id: 'filterNamaLembaga', dataColumn: 'Nama Lembaga', dependsOn: 'filterJenjang', dependencyColumn: '_filterJenjang' }
        ]
    };
  } catch (e) {
    return handleError('getPtkJumlahBulananPaudData', e);
  }
}

function getDaftarPtkPaudData() {
  try {
    const allData = getDataFromSheet('PTK_PAUD_DB');
    if (!allData || allData.length < 2) {
        return { headers: [], rows: [], filterConfigs: [] };
    }
    
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const jenjangIndex = headers.indexOf('Jenjang');
    const lembagaIndex = headers.indexOf('Nama Lembaga');
    const statusIndex = headers.indexOf('Status');
    
    const processedRows = dataRows.map(row => {
        const rowObject = {};
        headers.forEach((h, i) => rowObject[h] = row[i]);
        rowObject['_filterJenjang'] = row[jenjangIndex];
        rowObject['_filterNamaLembaga'] = row[lembagaIndex];
        rowObject['_filterStatus'] = row[statusIndex];
        return rowObject;
    });
    
    processedRows.sort((a,b) => (a['Nama'] || "").localeCompare(b['Nama'] || ""));

    return {
        headers: headers,
        rows: processedRows,
        filterConfigs: [
            { id: 'filterJenjang', dataColumn: '_filterJenjang' },
            { id: 'filterNamaLembaga', dataColumn: '_filterNamaLembaga', dependsOn: 'filterJenjang', dependencyColumn: '_filterJenjang' },
            { id: 'filterStatus', dataColumn: '_filterStatus' }
        ]
    };

  } catch (e) {
    return handleError('getDaftarPtkPaudData', e);
  }
}

function getKelolaPtkPaudData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    const allData = sheet.getDataRange().getValues();
    const headers = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    const indexedData = dataRows.map((row, index) => ({
      originalRowIndex: index + 2,
      rowData: row
    }));

    const parseDate = (value) => value instanceof Date && !isNaN(value) ? value.getTime() : 0;
    const updateIndex = headers.indexOf('Update');
    const dateInputIndex = headers.indexOf('Tanggal Input');
    
    indexedData.sort((a, b) => {
      const updateA = parseDate(a.rowData[updateIndex]);
      const updateB = parseDate(b.rowData[updateIndex]);
      if (updateB !== updateA) return updateB - updateA;
      const dateInputA = parseDate(a.rowData[dateInputIndex]);
      const dateInputB = parseDate(b.rowData[dateInputIndex]);
      return dateInputB - dateInputA;
    });

    const finalData = indexedData.map(item => {
      const rowDataObject = {
        _rowIndex: item.originalRowIndex,
        _source: 'PAUD',
      };
      headers.forEach((header, i) => {
        let cell = item.rowData[i];
        if (cell instanceof Date) {
          rowDataObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        } else {
          rowDataObject[header] = cell || "";
        }
      });
      return rowDataObject;
    });

    const displayHeaders = ["Nama", "Nama Lembaga", "Status", "NIP/NIY", "Jabatan", "Aksi"];
    return { headers: displayHeaders, rows: finalData };
  } catch (e) {
    return handleError('getKelolaPtkPaudData', e);
  }
}

function getPtkPaudDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      if (header === 'TMT' && values[i] instanceof Date) {
        rowData[header] = Utilities.formatDate(values[i], "UTC", "yyyy-MM-dd");
      } else {
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    return handleError('getPtkPaudDataByRow', e);
  }
}

function updatePtkPaudData(formData) {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    const rowIndex = formData.rowIndex;
    if (!rowIndex) throw new Error("Row index tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const oldValues = range.getValues()[0];

    formData['Update'] = new Date();

    const newRowValues = headers.map((header, index) => {
      return formData.hasOwnProperty(header) ? formData[header] : oldValues[index];
    });

    range.setValues([newRowValues]);
    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

function getNewPtkPaudOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.FORM_OPTIONS_DB.id);
    const sheet = ss.getSheetByName("Form PAUD");
    if (!sheet) throw new Error("Sheet 'Form PAUD' tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    const jenjangOptions = [...new Set(data.map(row => row[0]).filter(Boolean))].sort();
    const statusOptions = [...new Set(data.map(row => row[2]).filter(Boolean))].sort();
    const jabatanOptions = [...new Set(data.map(row => row[3]).filter(Boolean))].sort();
    
    const lembagaMap = {};
    data.forEach(row => {
      const jenjang = row[0];
      const lembaga = row[1];
      if (jenjang && lembaga) {
        if (!lembagaMap[jenjang]) lembagaMap[jenjang] = [];
        if (!lembagaMap[jenjang].includes(lembaga)) lembagaMap[jenjang].push(lembaga);
      }
    });
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
    return handleError('getNewPtkPaudOptions', e);
  }
}

function addNewPtkPaud(formData) {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(header => {
      if (header === 'Tanggal Input') return new Date();
      return formData[header] || "";
    });

    sheet.appendRow(newRow);
    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

function deletePtkPaudData(rowIndex, deleteCode) {
  try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");
    
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");
    
    const maxRows = sheet.getLastRow();
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > maxRows) throw new Error("Nomor baris tidak valid.");

    sheet.deleteRow(rowIndex);
    return "Data PTK berhasil dihapus.";
  } catch (e) {
    throw new Error(`Gagal menghapus data: ${e.message}`);
  }
}

function getPtkJumlahBulananSdData() {
  try {
    return getDataFromSheet('PTK_SD_JUMLAH_BULANAN');
  } catch (e) {
    return handleError('getPtkJumlahBulananSdData', e);
  }
}

function getDaftarPtkSdnData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_SD_DB;
    return SpreadsheetApp.openById(config.id).getSheetByName("PTK SDN").getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdnData', e);
  }
}

function getDaftarPtkSdsData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_SD_DB;
    return SpreadsheetApp.openById(config.id).getSheetByName("PTK SDS").getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdsData', e);
  }
}

function getKelolaPtkSdData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
    const sdnSheet = ss.getSheetByName("PTK SDN");
    const sdsSheet = ss.getSheetByName("PTK SDS");
    let combinedData = [];

    const processSheet = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      const data = sheet.getDataRange().getDisplayValues();
      const headers = data[0];
      const rows = data.slice(1);

      const idx = {
          unitKerja: headers.indexOf("Unit Kerja"),
          nama: headers.indexOf("Nama"),
          status: headers.indexOf("Status"),
          nipNiy: headers.indexOf("NIP/NIY"),
          jabatan: headers.indexOf("Jabatan"),
          tglInput: headers.indexOf("Input"),
          update: headers.indexOf("Update")
      };

      rows.forEach((row, index) => {
        if (!row[idx.nama]) return;
        
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
    
    return { rows: combinedData };

  } catch (e) {
    return handleError('getKelolaPtkSdData', e);
  }
}

function getNewPtkSdOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.FORM_OPTIONS_DB.id);
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

function addNewPtkSd(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
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
      if (header === 'NUPTK' && value) {
        return "'" + value;
      }
      return value;
    });

    sheet.appendRow(newRow);
    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

function getPtkSdDataByRow(rowIndex, source) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
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

function updatePtkSdData(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
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

    formData['Update'] = new Date();
    if (formData['NUPTK']) {
      formData['NUPTK'] = "'" + formData['NUPTK'];
    }

    const newRowValues = headers.map((header, index) => {
      return formData.hasOwnProperty(header) ? formData[header] : oldValues[index];
    });

    range.setValues([newRowValues]);
    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

function getKebutuhanPtkSdnData() {
  try {
    return getDataFromSheet('PTK_SD_KEBUTUHAN');
  } catch (e) {
    return handleError('getKebutuhanPtkSdnData', e);
  }
}

function deletePtkSdData(rowIndex, source, deleteCode) {
  try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
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
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > maxRows) throw new Error("Nomor baris tidak valid.");

    sheet.deleteRow(rowIndex);
    return "Data PTK berhasil dihapus.";
  } catch (e) {
    throw new Error(`Gagal menghapus data: ${e.message}`);
  }
}


/**
 * ===================================================================
 * ========================== MODUL: DATA SK =========================
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
 * [REFACTOR - FINAL V4] Mengambil data riwayat pengiriman SK.
 * Memperbaiki parsing tanggal untuk pengurutan yang benar.
 */
function getSKRiwayatData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.id)
                                .getSheetByName(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.sheet);

    const desiredHeaders = ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Tanggal SK", "Kriteria SK", "Dokumen", "Tanggal Unggah"];

    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: desiredHeaders, rows: [] };
    }
    
    const allData = sheet.getDataRange().getDisplayValues(); 
    const originalHeaders = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    // Fungsi parse tanggal yang lebih kuat untuk format "dd/MM/yyyy HH:mm:ss"
    const parseDate = (value) => {
        if (!value || typeof value !== 'string') return new Date(0);
        const parts = value.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
        if (parts) {
            // parts[2] - 1 karena bulan di JavaScript dimulai dari 0 (Januari=0)
            return new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5], parts[6]);
        }
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

    // Pengurutan sekarang akan berfungsi dengan benar
    structuredRows.sort((a, b) => {
      const dateB = parseDate(b['Tanggal Unggah']);
      const dateA = parseDate(a['Tanggal Unggah']);
      return dateB - dateA;
    });
    
    return {
      headers: desiredHeaders,
      rows: structuredRows
    };

  } catch (e) {
    return handleError('getSKRiwayatData', e);
  }
}

function getSKStatusData() {
  try {
    const data = getDataFromSheet('SK_BAGI_TUGAS'); 
    if (!data || data.length < 2) {
      return { headers: [], rows: [] };
    }

    const headers = data[0].map(h => String(h).trim());
    const dataRows = data.slice(1);

    const structuredRows = dataRows.map(row => {
      const rowObject = {};
      headers.forEach((header, index) => {
        rowObject[header] = row[index];
      });
      return rowObject;
    });
    
    return {
      headers: headers,
      rows: structuredRows
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
    
    const parseDate = (value) => value instanceof Date && !isNaN(value) ? value : new Date(0);

    const indexedData = dataRows.map((row, index) => ({
      row: row,
      originalIndex: index + 2
    }));
    
    const updateIndex = originalHeaders.indexOf('Update');
    const timestampIndex = originalHeaders.indexOf('Tanggal Unggah');

    indexedData.sort((a, b) => {
      const dateB_update = parseDate(b.row[updateIndex]);
      const dateA_update = parseDate(a.row[updateIndex]);
      if (dateB_update.getTime() !== dateA_update.getTime()) {
        return dateB_update - dateA_update;
      }
      const dateB_timestamp = parseDate(b.row[timestampIndex]);
      const dateA_timestamp = parseDate(a.row[timestampIndex]);
      return dateB_timestamp - dateA_timestamp;
    });

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
    
    // Ambil nilai mentah (RAW) untuk mendapatkan objek Date asli
    const rawValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Ambil nilai tampilan (DISPLAY) untuk konsistensi string/angka
    const displayValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const rowData = {};
    headers.forEach((header, i) => {
      // KUNCI PERBAIKAN: Format Tanggal SK ke YYYY-MM-DD
      if (header === 'Tanggal SK' && rawValues[i] instanceof Date) {
        // Format yang wajib untuk HTML input type="date"
        rowData[header] = Utilities.formatDate(rawValues[i], "UTC", "yyyy-MM-dd");
      } else {
        // Gunakan display value untuk field lain (Nomor SK, dll.)
        rowData[header] = displayValues[i];
      }
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
            const currentParentFolder = file.getParents().next();

            if (currentFileName !== newFilename || currentParentFolder.getName() !== newSemesterFolderName) {
                file.moveTo(newTargetFolder);
                file.setName(newFilename);
                fileUrl = file.getUrl();
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
    
    return "Data berhasil diperbarui!";
  } catch (e) {
    return handleError('updateSKData', e);
  }
}

function deleteSKData(rowIndex, deleteCode) {
  try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

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
 * ======================== MODUL: DATA SIABA ========================
 * ===================================================================
 */

function getSiabaFilterOptions() {
  try {
    const ssDropdown = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
    const sheetUnitKerja = ssDropdown.getSheetByName("Unit Siaba");
    let unitKerjaOptions = [];
    if (sheetUnitKerja && sheetUnitKerja.getLastRow() > 1) {
      unitKerjaOptions = sheetUnitKerja.getRange(2, 1, sheetUnitKerja.getLastRow() - 1, 1)
                                      .getDisplayValues().flat().filter(Boolean).sort();
    }

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

function getSiabaPresensiData(filters) {
  try {
    const { tahun, bulan } = filters;
    if (!tahun || !bulan) throw new Error("Filter Tahun dan Bulan wajib diisi.");

    const config = SPREADSHEET_CONFIG.SIABA_REKAP;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => String(row[0]) === String(tahun) && String(row[1]) === String(bulan));

    const startIndex = 2;
    const endIndex = 86;

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => row.slice(startIndex, endIndex + 1));
    
    const tpIndex = displayHeaders.indexOf('TP');
    const taIndex = displayHeaders.indexOf('TA');
    const plaIndex = displayHeaders.indexOf('PLA');
    const tapIndex = displayHeaders.indexOf('TAp');
    const tuIndex = displayHeaders.indexOf('TU');
    const namaIndex = displayHeaders.indexOf('Nama');

    if (tpIndex !== -1) {
      displayRows.sort((a, b) => {
        const compareDesc = (index) => {
            if (index === -1) return 0;
            const valB = parseInt(b[index], 10) || 0;
            const valA = parseInt(a[index], 10) || 0;
            return valB - valA;
        };
        
        let diff = compareDesc(tpIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(taIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(plaIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(tapIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(tuIndex);
        if (diff !== 0) return diff;

        if (namaIndex !== -1) {
            return (a[namaIndex] || "").localeCompare(b[namaIndex] || "");
        }
        return 0;
      });
    }

    return { headers: displayHeaders, rows: displayRows };

  } catch (e) {
    return handleError('getSiabaPresensiData', e);
  }
}

function getSiabaTidakPresensiFilterOptions() {
  try {
    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
         throw new Error("Sheet 'Rekap Script' untuk data Tidak Presensi tidak ditemukan atau kosong.");
    }

    const filterData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
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

function getSiabaTidakPresensiData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    if (!tahun || !bulan) throw new Error("Filter Tahun dan Bulan wajib diisi.");

    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => {
      const tahunMatch = String(row[0]) === String(tahun);
      const bulanMatch = String(row[1]) === String(bulan);
      const unitKerjaMatch = (unitKerja === "Semua") || (String(row[2]) === String(unitKerja));
      return tahunMatch && bulanMatch && unitKerjaMatch;
    });

    const startIndex = 3;
    const endIndex = 8;

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => row.slice(startIndex, endIndex + 1));
    
    const jumlahIndex = displayHeaders.indexOf('Jumlah');
    const namaIndex = displayHeaders.indexOf('Nama');

    displayRows.sort((a, b) => {
      const valB_jumlah = (jumlahIndex !== -1) ? (parseInt(b[jumlahIndex], 10) || 0) : 0;
      const valA_jumlah = (jumlahIndex !== -1) ? (parseInt(a[jumlahIndex], 10) || 0) : 0;
      if (valB_jumlah !== valA_jumlah) {
        return valB_jumlah - valA_jumlah;
      }
      if (namaIndex !== -1) {
          return (a[namaIndex] || "").localeCompare(b[namaIndex] || "");
      }
      return 0;
    });

    return { headers: displayHeaders, rows: displayRows };
  } catch (e) {
    return handleError('getSiabaTidakPresensiData', e);
  }
}
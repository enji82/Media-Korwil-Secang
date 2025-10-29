// File: Config.gs

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
  PTK_PAUD_JUMLAH_BULANAN: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Data PAUD Bulanan" },
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
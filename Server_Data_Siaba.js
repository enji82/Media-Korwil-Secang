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
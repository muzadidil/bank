function jalankanSetupMassal() {
  // Daftar 12 File Spreadsheet dan Nama Tab Utamanya
  const daftarFile = [
    { id: "1e8Og97M5gTTlR7NUWk6YDLocZEZJg6ypCFLp_2mtl6w", bulan: "JAN_2023" },
    { id: "1Usdw8EMDpJBlDv1zkcteoj8tEgRmocNHuPmOeSm0s5Q", bulan: "FEB_2023" },
    { id: "14VmCTjGRZ4jkWWz2LcnttoLb-kaUPgo86bjOezQBlV4", bulan: "MAR_2023" },
    { id: "1k6aOukMobmd-kJdoNI1uMNLSTH67-63xOByr2mlT7To", bulan: "APR_2023" },
    { id: "1--0ZS9CxyYGA0moxG8xUiywvORuaMLbBDv3w3aSx-Og", bulan: "MEI_2023" },
    { id: "1JjCM1NiVYyoRVdG7XrbdLV5j4TxqhRcJ5dfISwsJbM4", bulan: "JUN_2023" },
    { id: "1HrDJRSKH7asVVlOGn4Op1MH4Rr8itXoJpVId6WGUERA", bulan: "JUL_2023" },
    { id: "11NzCKoEjr0JvWkTopOqkOxqi8yw3M1aAKzPnbsX3nfo", bulan: "AGU_2023" },
    { id: "1cA4rBuF5tSFYz1SbMBcJlIZBZ2Iy75jK3A9WReeHZew", bulan: "SEP_2023" },
    { id: "1kLaUw3mazHJr0c-BcP7JXPVSKmx54dH8JfgbwYZDaNo", bulan: "OKT_2023" },
    { id: "1W1z_Jz1FMdBB4DbqLI3QQfmLHBAAA1NEA2-iarT1F-c", bulan: "NOV_2023" },
    { id: "1jLLM7RowyAlqW8_utve-PxKKawTbcRuWVtKmI2wg8Ww", bulan: "DES_2023" }
  ];

  // Daftar Kategori Sheet Baru
  const daftarKategori = ["EXSPEDISI", "V_SALES", "OPERASIONAL", "NURUL_AINI"];

  Logger.log("Memulai proses pembuatan sheet massal...");

  // Looping ke setiap file
  daftarFile.forEach(file => {
    try {
      const ss = SpreadsheetApp.openById(file.id);
      
      // Looping ke setiap kategori untuk file tersebut
      daftarKategori.forEach(kategori => {
        let sheetBaru = ss.getSheetByName(kategori);
        
        // Jika sheet belum ada, buat baru
        if (!sheetBaru) {
          sheetBaru = ss.insertSheet(kategori);
        }
        
        // Buat rumus secara dinamis sesuai nama bulan di file ini
        const rumusA1 = `=IFERROR(QUERY(${file.bulan}!A:I, "SELECT * WHERE D IS NOT NULL AND UPPER(I) = '${kategori}'", 1), "Data ${kategori} belum ditemukan")`;
        const rumusJ1 = `=SUM(D1:D)`;
        
        // Tanam rumus ke cell A1 dan J1
        sheetBaru.getRange("A1").setFormula(rumusA1);
        sheetBaru.getRange("J1").setFormula(rumusJ1);
      });
      
      Logger.log("✅ SUKSES: File " + file.bulan + " sudah diperbarui.");
    } catch (e) {
      Logger.log("❌ ERROR pada " + file.bulan + " (ID: " + file.id + "): " + e.message);
    }
  });

  Logger.log("🎉 SEMUA PROSES SELESAI!");
}

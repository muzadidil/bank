/**
 * ZASHA MUTASI - PROJECT PRODUKSI
 */

function doGet(e) {
  // Default ke JAN_2023 jika parameter bulan tidak ada
  var bulan = e.parameter.bulan || "JAN_2023"; 
  
  var template = HtmlService.createTemplateFromFile('View_Produksi');
  template.bulanAktif = bulan; 
  
  return template.evaluate()
      .setTitle('Laporan Produksi - ' + bulan)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheetData(namaBulan) {
  const sheetId = "1e8Og97M5gTTlR7NUWk6YDLocZEZJg6ypCFLp_2mtl6w"; 
  const ss = SpreadsheetApp.openById(sheetId);
  
  // 1. Ambil Data Transaksi dari Sheet Master (Bulan terkait)
  const masterSheet = ss.getSheetByName(namaBulan); 
  if (!masterSheet) return { error: "Sheet Master " + namaBulan + " tidak ditemukan!" };
  const values = masterSheet.getDataRange().getValues();

  // 2. Ambil Angka Total Khusus dari Sheet "PRODUKSI" Sel J1
  const produksiSheet = ss.getSheetByName("PRODUKSI");
  const totalProduksiJ1 = produksiSheet ? produksiSheet.getRange("J1").getValue() : 0;

  const filteredData = [];
  for (let i = 1; i < values.length; i++) {
    let nominal = values[i][3]; // Kolom D
    let kategori = values[i][8] ? values[i][8].toString().toUpperCase() : ""; // Kolom I
    
    // FILTER: Hanya tampilkan jika kategori adalah PRODUKSI
    if (nominal && nominal !== "" && kategori === "PRODUKSI") {
      filteredData.push({
        baris: i + 1,
        tanggal: formatDate(values[i][0]),
        keterangan: values[i][2],
        nominal: nominal,
        kategori: kategori
      });
    }
  }

  return {
    transactions: filteredData,
    totalMaster: totalProduksiJ1 
  };
}

function updateKategori(baris, kategoriBaru, namaBulan) {
  try {
    const sheetId = "1e8Og97M5gTTlR7NUWk6YDLocZEZJg6ypCFLp_2mtl6w";
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(namaBulan);
    sheet.getRange(baris, 9).setValue(kategoriBaru);
    return "Sukses";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM");
  }
  return date;
}

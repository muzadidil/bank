/**
 * ZASHA MUTASI - FINAL ENGINE (STANDALONE & STABLE)
 * ID: 14VmCTjGRZ4jkWWz2LcnttoLb-kaUPgo86bjOezQBlV4
 */

const SPREADSHEET_ID = "14VmCTjGRZ4jkWWz2LcnttoLb-kaUPgo86bjOezQBlV4";

function doGet(e) {
  try {
    var page = e.parameter.p || "bank"; 
    var bulan = e.parameter.bulan || "MAR_2023";
    var tmp = HtmlService.createTemplateFromFile('Index');
    
    tmp.url = ScriptApp.getService().getUrl(); 
    tmp.page = page;
    tmp.bulan = bulan;
    
    const judulMap = {
      "bank": "Informasi Mutasi",
      "pribadi": "Laporan Pribadi",
      "produksi": "Laporan Produksi",
      "kosong": "Kategori Kosong",
      "exspedisi": "Kategori Exspedisi",
      "v_sales": "Kategori V_Sales",
      "operasional": "Kategori Operasional",
      "nurul_aini": "Kategori Nurul Aini"
    };
    tmp.judul = judulMap[page] || "Dashboard Zasha";

    return tmp.evaluate()
        .setTitle(tmp.judul + " - " + bulan)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("<b>Fatal Error:</b> " + err.message);
  }
}

function getDataFromSheet(page, bulan) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const masterSheet = ss.getSheetByName(bulan);
    
    if (!masterSheet) return { error: "Sheet '" + bulan + "' tidak ditemukan!", transactions: [], categories: [] };
    
    const allValues = masterSheet.getDataRange().getValues();
    const categories = [...new Set(allValues.slice(1).map(row => row[8]).filter(val => val !== ""))];

    if (page === "pribadi") return { ...getPribadiData(ss, bulan), categories };
    if (page === "produksi") return { ...getProduksiData(ss, bulan), categories };
    if (page === "kosong") return { ...getKosongData(ss, bulan), categories };
    
    // JALUR YANG DIPERBAIKI: Mengarah ke fungsinya masing-masing
    if (page === "exspedisi") return { ...getExspedisiData(ss, bulan), categories };
    if (page === "v_sales") return { ...getV_salesData(ss, bulan), categories };
    if (page === "operasional") return { ...getOperasionalData(ss, bulan), categories };
    if (page === "nurul_aini") return { ...getNurul_ainiData(ss, bulan), categories };
    
    return { data: getBankData(ss, bulan), categories };
  } catch (e) {
    return { error: "Server Error: " + e.toString() };
  }
}

function getBankData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  return values.slice(1).map(row => ({
    tanggal: formatDate(row[0]), keterangan: row[1], debet: row[3], kredit: row[4], saldo: row[6]
  }));
}

function getPribadiData(ss, namaBulan) {
  const sheetRef = ss.getSheetByName("PRIBADI");
  const total = sheetRef ? sheetRef.getRange("J1").getValue() : 0;
  const masterSheet = ss.getSheetByName(namaBulan);
  const values = masterSheet.getDataRange().getValues();
  const filtered = values.slice(1).filter(row => row[3] && (row[8]||"").toString().toUpperCase() === "PRIBADI").map(row => ({
    baris: values.indexOf(row) + 1, tanggal: formatDate(row[0]), keterangan: row[2], nominal: row[3]
  }));
  return { transactions: filtered, total: total };
}

function getProduksiData(ss, namaBulan) {
  const sheetRef = ss.getSheetByName("PRODUKSI");
  const total = sheetRef ? sheetRef.getRange("J1").getValue() : 0;
  const masterSheet = ss.getSheetByName(namaBulan);
  const values = masterSheet.getDataRange().getValues();
  const filtered = values.slice(1).filter(row => row[3] && (row[8]||"").toString().toUpperCase() === "PRODUKSI").map(row => ({
    baris: values.indexOf(row) + 1, tanggal: formatDate(row[0]), keterangan: row[2], nominal: row[3]
  }));
  return { transactions: filtered, total: total };
}

function getKosongData(ss, namaBulan) {
  const sheetRef = ss.getSheetByName("KOSONG");
  const total = sheetRef ? sheetRef.getRange("J1").getValue() : 0;
  const masterSheet = ss.getSheetByName(namaBulan);
  const values = masterSheet.getDataRange().getValues();
  const filtered = values.slice(1).filter(row => row[3] && (row[8]||"").toString().trim() === "").map(row => ({
    baris: values.indexOf(row) + 1, tanggal: formatDate(row[0]), keterangan: row[2], nominal: row[3]
  }));
  return { transactions: filtered, total: total };
}

// ==========================================
// FUNGSI-FUNGSI BARU ZASHA ONLINE
// ==========================================

function getExspedisiData(ss, namaBulan) {
  const sheetRef = ss.getSheetByName("EXSPEDISI");
  const total = sheetRef ? sheetRef.getRange("J1").getValue() : 0;
  const masterSheet = ss.getSheetByName(namaBulan);
  const values = masterSheet.getDataRange().getValues();
  const filtered = values.slice(1).filter(row => row[3] && (row[8]||"").toString().toUpperCase() === "EXSPEDISI").map(row => ({
    baris: values.indexOf(row) + 1, tanggal: formatDate(row[0]), keterangan: row[2], nominal: row[3]
  }));
  return { transactions: filtered, total: total };
}

function getV_salesData(ss, namaBulan) {
  const sheetRef = ss.getSheetByName("V_SALES");
  const total = sheetRef ? sheetRef.getRange("J1").getValue() : 0;
  const masterSheet = ss.getSheetByName(namaBulan);
  const values = masterSheet.getDataRange().getValues();
  const filtered = values.slice(1).filter(row => row[3] && (row[8]||"").toString().toUpperCase() === "V_SALES").map(row => ({
    baris: values.indexOf(row) + 1, tanggal: formatDate(row[0]), keterangan: row[2], nominal: row[3]
  }));
  return { transactions: filtered, total: total };
}

function getOperasionalData(ss, namaBulan) {
  const sheetRef = ss.getSheetByName("OPERASIONAL");
  const total = sheetRef ? sheetRef.getRange("J1").getValue() : 0;
  const masterSheet = ss.getSheetByName(namaBulan);
  const values = masterSheet.getDataRange().getValues();
  const filtered = values.slice(1).filter(row => row[3] && (row[8]||"").toString().toUpperCase() === "OPERASIONAL").map(row => ({
    baris: values.indexOf(row) + 1, tanggal: formatDate(row[0]), keterangan: row[2], nominal: row[3]
  }));
  return { transactions: filtered, total: total };
}

function getNurul_ainiData(ss, namaBulan) {
  const sheetRef = ss.getSheetByName("NURUL_AINI");
  const total = sheetRef ? sheetRef.getRange("J1").getValue() : 0;
  const masterSheet = ss.getSheetByName(namaBulan);
  const values = masterSheet.getDataRange().getValues();
  const filtered = values.slice(1).filter(row => row[3] && (row[8]||"").toString().toUpperCase() === "NURUL_AINI").map(row => ({
    baris: values.indexOf(row) + 1, tanggal: formatDate(row[0]), keterangan: row[2], nominal: row[3]
  }));
  return { transactions: filtered, total: total };
}

// ==========================================

function updateBatchKategori(arrBaris, kategoriBaru, namaBulan, pageType) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetMaster = ss.getSheetByName(namaBulan);
    arrBaris.forEach(baris => {
      sheetMaster.getRange(parseInt(baris), 9).setValue(kategoriBaru);
    });
    
    const sheetTujuan = ss.getSheetByName(pageType.toUpperCase());
    let newTotal = sheetTujuan ? sheetTujuan.getRange("J1").getValue() : 0;

    return { message: "Data berhasil dipindahkan", newTotal: newTotal };
  } catch (e) { return { error: e.toString() }; }
}

function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM");
  }
  return (date || "").toString();
}
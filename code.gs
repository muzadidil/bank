function doGet(e) {
  // Mengambil nama sheet dari URL parameter (?sheet=NAMA_SHEET)
  // Jika tidak ada parameter, otomatis default ke JAN_2023
  var sheetParam = e.parameter.sheet || "JAN_2023"; 
  
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Kirim variabel ke Index.html
  template.namaSheetDinamis = sheetParam;
  template.judulTampilan = sheetParam.replace("_", " "); // Contoh: FEB_2023 jadi FEB 2023
  
  return template.evaluate()
      .setTitle('Mutasi Rekening ' + sheetParam)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  // Jika sheet tidak ditemukan, kirim pesan error
  if (!sheet) return "Error: Sheet [" + sheetName + "] tidak ditemukan!";

  const range = sheet.getDataRange();
  const values = range.getValues();
  
  // Mapping Kolom: A=0(Tgl), C=2(Ket), D=3(Db), E=4(Cr), G=6(Saldo)
  const filteredData = values.slice(1).map(row => {
    if (!row[2]) return null; // Abaikan baris kosong

    return {
      tanggal: formatDate(row[0]),
      keterangan: row[2],
      debet: row[3],
      kredit: row[4],
      saldo: row[6]
    };
  }).filter(item => item !== null);

  return filteredData;
}

function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM");
  }
  return date;
}

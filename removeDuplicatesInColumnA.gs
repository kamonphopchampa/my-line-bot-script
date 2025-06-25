function removeDuplicatesInColumnG() {
  // ใช้ Sheet ID ที่กำหนด
  var sheetId = "xxxxxxxxxxx";
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadsheet.getSheetByName("USER ID");
  
  if (!sheet) {
    Logger.log("Sheet 'USER ID' not found in the specified spreadsheet.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var seen = new Set();
  var rowsToDelete = [];

  for (var i = 1; i < data.length; i++) {  // เริ่มจาก index 1 เพื่อข้ามหัวตาราง
    var value = data[i][6]; // คอลัมน์ G คือ index 6 (เริ่มจาก 0)
    
    if (value && seen.has(value)) {
      rowsToDelete.push(i + 1); // เก็บแถวที่ต้องลบ (1-based index)
      Logger.log("Found duplicate in row " + (i + 1) + ": " + value);
    } else if (value) { // เพิ่มเงื่อนไข value ไม่เป็นค่าว่าง
      seen.add(value);
    }
  }

  // ลบจากด้านล่างขึ้นบนเพื่อไม่ให้ index เปลี่ยน
  if (rowsToDelete.length > 0) {
    Logger.log("Deleting " + rowsToDelete.length + " duplicate rows: " + rowsToDelete.join(", "));
    rowsToDelete.reverse().forEach(row => sheet.deleteRow(row));
    Logger.log("Duplicate removal completed.");
  } else {
    Logger.log("No duplicates found in column G.");
  }
}

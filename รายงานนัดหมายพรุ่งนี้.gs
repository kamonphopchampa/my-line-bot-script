function รายงานนัดหมายพรุ่งนี้() {
  // ใช้ Sheet ID ที่กำหนด
  const SHEET_ID = 'xxxxxxxxxx';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  const monthNames = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
                      "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthName = monthNames[tomorrow.getMonth()];
  const sheetName = monthName;
  const reportSheetName = "รายงาน" + monthName;

  const sourceSheet = ss.getSheetByName(sheetName);
  const reportSheet = ss.getSheetByName(reportSheetName);

  if (!sourceSheet || !reportSheet) {
    Logger.log("ไม่พบชีต: " + sheetName + " หรือ " + reportSheetName);
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const formulaOutput = [];
  const todayDateOutput = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateCell = row[0];

    if (dateCell instanceof Date && isSameDate(dateCell, tomorrow)) {
      const formula = `='${sheetName}'!D${i + 1}&" HN: "&'${sheetName}'!C${i + 1}&" นัดหมายวันที่: "&TEXT('${sheetName}'!A${i + 1},"d mmmm yyyy")&" เวลา: "&'${sheetName}'!B${i + 1}` +
        `&" เบอร์โทร: "&'${sheetName}'!F${i + 1}&" รายละเอียด: "&'${sheetName}'!E${i + 1}`;
      formulaOutput.push([formula]);

      // const formula = `='${sheetName}'!D${i + 1}&" HN: "&'${sheetName}'!C${i + 1}&" นัดหมายวันที่: "&TEXT('${sheetName}'!A${i + 1},"d mmmm yyyy")&" เวลา: "&'${sheetName}'!B${i + 1}` +
      //   `&" เบอร์โทร: "&'${sheetName}'!G${i + 1}&" - "&'${sheetName}'!F${i + 1}&" รายละเอียด: "&'${sheetName}'!E${i + 1}`;
      // formulaOutput.push([formula]);
      // todayDateOutput.push([today]); // เพิ่มวันที่วันนี้ในคอลัมน์ I
    }
  }

  // ล้างข้อมูลเก่าในคอลัมน์ B และ I
  reportSheet.getRange("B2:B").clearContent();
  reportSheet.getRange("I2:I").clearContent();

  if (formulaOutput.length > 0) {
    // ลงสูตรในคอลัมน์ B
    reportSheet.getRange(2, 2, formulaOutput.length, 1).setFormulas(formulaOutput);
    // ลงวันที่วันนี้ในคอลัมน์ I
    // reportSheet.getRange(2, 9, todayDateOutput.length, 1).setValues(todayDateOutput);
    Logger.log("พบข้อมูลนัดหมายพรุ่งนี้ " + formulaOutput.length + " รายการ ในชีต " + sheetName);
  } else {
    Logger.log("ไม่มีข้อมูลที่ตรงกับพรุ่งนี้ในชีต " + sheetName);
  }
}

function isSameDate(d1, d2) {
  const y1 = d1.getFullYear() > 2500 ? d1.getFullYear() - 543 : d1.getFullYear();
  const y2 = d2.getFullYear();
  return d1.getDate() === d2.getDate() &&
         d1.getMonth() === d2.getMonth() &&
         y1 === y2;
}

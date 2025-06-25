// ฟังก์ชันสำหรับตั้ง Trigger อัตโนมัติ (รันทุกวันเวลา 08:00)
function createDailyTrigger2() {
  ScriptApp.newTrigger('sendMedicalReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8) // 8 โมงเช้า
    .create();
  
  Logger.log("สร้าง Daily Trigger สำเร็จ - จะทำงานทุกวันเวลา 08:00");
}/**
 * แจ้งเตือนนัดหมาย.gs - ระบบแจ้งเตือนนัดหมายแพทย์
 * อ้างอิงตัวแปรจากไฟล์ "ส่งไลน์.gs":
 * - SHEET_ID
 * - CHANNEL_ACCESS_TOKEN  
 * - MONTH_NAMES_TH
 */

function sendMedicalReminders1() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const today = new Date();
  
  // วนลูปตรวจสอบทุกเดือน
  for (let monthIndex = 0; monthIndex < 12; monthIndex++) {
    const monthName = MONTH_NAMES_TH[monthIndex];
    const sheetName = monthName; // หรือ "รายงาน" + monthName ตามที่คุณใช้
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("ไม่พบชีท: " + sheetName);
      continue;
    }
    
    Logger.log("พบชีท: " + sheetName + " - กำลังตรวจสอบ...");
    checkAppointmentReminders(sheet, today);
  }
}

function checkAppointmentReminders(sheet, today) {
  // ตรวจสอบว่า sheet มีค่าหรือไม่
  if (!sheet) {
    Logger.log("Sheet is null or undefined");
    return;
  }

  const lastRow = sheet.getLastRow();
  Logger.log(`ชีท ${sheet.getName()} มีข้อมูล ${lastRow} แถว`);
  
  if (lastRow < 2) {
    Logger.log("ไม่มีข้อมูลในชีท " + sheet.getName());
    return; // ไม่มีข้อมูล
  }
  
  // ดึงข้อมูลทั้งหมดที่จำเป็น
  const dataRange = sheet.getRange("A2:I" + lastRow);
  const data = dataRange.getValues();
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const appointmentDate = row[0]; // คอลัมน์ A - วันที่นัดหมาย
    const time = row[1];           // คอลัมน์ B - เวลา
    const hn = row[2];             // คอลัมน์ C - HN
    const patientName = row[3];    // คอลัมน์ D - ชื่อ
    const details = row[4];        // คอลัมน์ E - รายละเอียด
    const phone = row[5];          // คอลัมน์ F - เบอร์โทร
    const userId = row[6];         // คอลัมน์ G - USER ID
    const notes = row[7];          // คอลัมน์ H - หมายเหตุ
    const reminderDate = row[8];   // คอลัมน์ I - วันที่โทรนัด/แจ้งเตือน
    
    // ตรวจสอบว่าข้อมูลครบถ้วนหรือไม่
    if (!appointmentDate || !patientName || !reminderDate) {
      Logger.log(`ข้ามแถว ${i+2}: ข้อมูลพื้นฐานไม่ครบ - วันที่นัด:${appointmentDate}, ชื่อ:${patientName}, วันแจ้งเตือน:${reminderDate}`);
      continue;
    }

    // ตรวจสอบ User ID แยกต่างหาก เพื่อให้เห็นปัญหาชัดเจน
    const userIdStr = userId ? userId.toString().trim() : "";
    if (!userIdStr || userIdStr === "") {
      Logger.log(`⚠️  แถว ${i+2} (${patientName}): ไม่มี User ID - ข้ามการส่งข้อความ`);
      // อย่า continue เพื่อให้เห็นการเปรียบเทียบวันที่
    }
    
    // ตรวจสอบว่าวันนี้ตรงกับวันที่ต้องแจ้งเตือนหรือไม่
    const reminderDateObj = new Date(reminderDate);
    const todayString = formatDateForComparison(today);
    const reminderDateString = formatDateForComparison(reminderDateObj);
    
    Logger.log(`🔍 ตรวจสอบแถว ${i+2} (${patientName}): วันนี้=${todayString}, วันแจ้งเตือน=${reminderDateString}, User ID=${userIdStr || 'ไม่มี'}`);
    
    if (todayString === reminderDateString) {
      if (userIdStr && userIdStr !== "") {
        Logger.log(`✅ ส่งการแจ้งเตือนให้ ${patientName} (${userIdStr})`);
        sendAppointmentReminder(userIdStr, appointmentDate, time, patientName, details, phone, hn);
        Utilities.sleep(500); // ป้องกัน rate limit
      } else {
        Logger.log(`❌ วันที่ตรงกัน แต่ไม่มี User ID สำหรับ ${patientName}`);
      }
    }
  }
}

function formatDateForComparison(date) {
  // ตรวจสอบว่าเป็น Date object ที่ถูกต้องหรือไม่
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    Logger.log(`❌ วันที่ไม่ถูกต้อง: ${date}`);
    return null;
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  // แปลงปี พ.ศ. เป็น ค.ศ. ถ้าจำเป็น
  const adjustedYear = year > 2500 ? year - 543 : year;
  
  const result = `${adjustedYear}-${month}-${day}`;
  Logger.log(`📅 แปลงวันที่: ${date} -> ${result}`);
  return result;
}

function sendAppointmentReminder(userId, appointmentDate, time, patientName, details, phone, hn) {
  // ตรวจสอบ userId อีกครั้งก่อนส่ง
  if (!userId || userId.toString().trim() === "") {
    Logger.log(`❌ ไม่สามารถส่งข้อความได้ - User ID ว่างเปล่าสำหรับ ${patientName}`);
    return;
  }

  const appointmentDateObj = new Date(appointmentDate);
  const formattedDate = formatThaiDate(appointmentDateObj);
  const formattedTime = time ? formatTime(time) : "ไม่ระบุเวลา";
  
  Logger.log(`📤 กำลังส่งข้อความให้ User ID: ${userId}, ชื่อ: ${patientName}`);
  
  const flexMessage = {
    "type": "flex",
    "altText": `แจ้งเตือนนัดหมาย: ${patientName}`,
    "contents": {
      "type": "bubble",
      "size": "giga",
      "header": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "🏥 แจ้งเตือนนัดหมายแพทย์",
            "weight": "bold",
            "size": "lg",
            "color": "#ffffff"
          }
        ],
        "backgroundColor": "#42C2FF",
        "paddingAll": "lg"
      },
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "text",
                "text": `คุณ${patientName}`,
                "size": "xl",
                "weight": "bold",
                "color": "#333333"
              },
              {
                "type": "text",
                "text": "ใกล้ถึงวันที่หมอนัดพบแพทย์แล้ว",
                "size": "sm",
                "color": "#666666",
                "margin": "xs"
              }
            ],
            "margin": "none"
          },
          {
            "type": "separator",
            "margin": "lg"
          },
          {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  {
                    "type": "text",
                    "text": "📅 วันที่นัด:",
                    "size": "sm",
                    "weight": "bold",
                    "color": "#333333",
                    "flex": 2
                  },
                  {
                    "type": "text",
                    "text": `พรุ่งนี้ (${formattedDate})`,
                    "size": "sm",
                    "color": "#FF5551",
                    "weight": "bold",
                    "flex": 3
                  }
                ],
                "margin": "sm"
              },
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  {
                    "type": "text",
                    "text": "⏰ เวลา:",
                    "size": "sm",
                    "weight": "bold",
                    "color": "#333333",
                    "flex": 2
                  },
                  {
                    "type": "text",
                    "text": formattedTime,
                    "size": "sm",
                    "color": "#666666",
                    "flex": 3
                  }
                ],
                "margin": "sm"
              },
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  {
                    "type": "text",
                    "text": "🩺 วัตถุประสงค์:",
                    "size": "sm",
                    "weight": "bold",
                    "color": "#333333",
                    "flex": 2
                  },
                  {
                    "type": "text",
                    "text": details || "ตรวจสุขภาพ",
                    "size": "sm",
                    "color": "#666666",
                    "wrap": true,
                    "flex": 3
                  }
                ],
                "margin": "sm"
              }
            ],
            "margin": "lg"
          }
        ],
        "paddingAll": "lg"
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": `HN: ${hn || 'ไม่ระบุ'}`,
                "size": "xs",
                "color": "#999999",
                "flex": 1
              },
              {
                "type": "text",
                "text": `Tel: ${phone || 'ไม่ระบุ'}`,
                "size": "xs",
                "color": "#999999",
                "flex": 1,
                "align": "end"
              }
            ]
          },
          {
            "type": "separator",
            "margin": "sm"
          },
          {
            "type": "text",
            "text": "💊 อย่าลืมมาตรงเวลานะครับ",
            "size": "sm",
            "color": "#42C2FF",
            "weight": "bold",
            "align": "center",
            "margin": "sm"
          }
        ],
        "paddingAll": "lg"
      }
    }
  };

  const payload = {
    "to": userId,
    "messages": [flexMessage]
  };

  const options = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN
    },
    "payload": JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
    Logger.log(`ส่งข้อความสำเร็จให้ ${patientName}: ${response.getResponseCode()}`);
  } catch (error) {
    Logger.log(`เกิดข้อผิดพลาดในการส่งข้อความให้ ${patientName}: ${error.toString()}`);
  }
}

function formatThaiDate2(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return "วันที่ไม่ถูกต้อง";
  }

  const day = date.getDate();
  const month = MONTH_NAMES_TH[date.getMonth()];
  
  // ตรวจสอบปีและแปลงให้ถูกต้อง
  let year = date.getFullYear();
  
  // ถ้าปีเป็น พ.ศ. อยู่แล้ว (มากกว่า 2500) ให้แปลงเป็น ค.ศ. ก่อน
  if (year > 2500) {
    year = year - 543;
  }
  
  // จากนั้นแปลงเป็น พ.ศ. สำหรับแสดงผล
  const buddhistYear = year + 543;
  
  Logger.log(`📅 แปลงวันที่สำหรับแสดงผล: ${date} -> ${day} ${month} ${buddhistYear}`);
  
  return `${day} ${month} ${buddhistYear}`;
}

function formatTime(time) {
  if (!time) return "ไม่ระบุเวลา";
  
  // ถ้า time เป็น Date object
  if (time instanceof Date) {
    const hours = String(time.getHours()).padStart(2, '0');
    const minutes = String(time.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes} น.`;
  }
  
  // ถ้า time เป็น string
  return time.toString() + (time.toString().includes('น.') ? '' : ' น.');
}

// ฟังก์ชันสำหรับ test การทำงาน - ตรวจสอบข้อมูลในชีต
function testMedicalReminders() {
  Logger.log("🔍 เริ่มทดสอบระบบแจ้งเตือนนัดหมาย...");
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();
  
  Logger.log("📋 ชีตทั้งหมดในไฟล์:");
  sheets.forEach(sheet => {
    Logger.log("- " + sheet.getName());
  });
  
  // ทดสอบดูข้อมูลในชีตแรกที่พบ
  if (sheets.length > 0) {
    const testSheet = sheets[0];
    Logger.log(`🔬 ตรวจสอบข้อมูลในชีต: ${testSheet.getName()}`);
    
    const lastRow = testSheet.getLastRow();
    if (lastRow >= 2) {
      const sampleData = testSheet.getRange("A2:I2").getValues()[0];
      Logger.log("📊 ข้อมูลตัวอย่างแถวที่ 2:");
      Logger.log(`A (วันที่นัด): ${sampleData[0]}`);
      Logger.log(`B (เวลา): ${sampleData[1]}`);
      Logger.log(`C (HN): ${sampleData[2]}`);
      Logger.log(`D (ชื่อ): ${sampleData[3]}`);
      Logger.log(`E (รายละเอียด): ${sampleData[4]}`);
      Logger.log(`F (เบอร์โทร): ${sampleData[5]}`);
      Logger.log(`G (User ID): ${sampleData[6]}`);
      Logger.log(`H (หมายเหตุ): ${sampleData[7]}`);
      Logger.log(`I (วันแจ้งเตือน): ${sampleData[8]}`);
    }
  }
  
  Logger.log("🚀 เริ่มการตรวจสอบจริง...");
  sendMedicalReminders();
  Logger.log("✅ ทดสอบเสร็จสิ้น");
}

// ฟังก์ชันสำหรับเพิ่มข้อมูล Test User ID (สำหรับทดสอบ)
function addTestUserId() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("มิถุนายน"); // ชีตที่มีข้อมูล
  
  if (!sheet) {
    Logger.log("❌ ไม่พบชีต มิถุนายน");
    return;
  }
  
  // ใส่ User ID ทดสอบในแถวที่ 6 (สมชาย) เพื่อทดสอบ
  const testUserId = "Ub7fb81c85b0bb6c8be5bbafdeeb7fb3b"; // ใส่ User ID จริงของคุณ
  sheet.getRange("G6").setValue(testUserId);
  
  Logger.log(`✅ เพิ่ม Test User ID สำหรับ สมชาย ในแถว 6: ${testUserId}`);
  Logger.log("💡 ตอนนี้สามารถทดสอบการส่งข้อความได้แล้ว");
}

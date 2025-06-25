/**
 * Line to Calendar.gs - ระบบบันทึกข้อมูลนัดหมายจาก LINE ไป Google Calendar
 * หมายเหตุ: ใช้ตัวแปร Constants จากไฟล์ "ส่งไลน์.gs"
 */

// กำหนดชื่อเดือนภาษาไทย (เผื่อไม่มีใน Constants)
const MONTH_NAMES_TH_LOCAL = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// ใช้ ACCESS_TOKEN ที่ถูกต้องจากไฟล์ "ส่งไลน์.gs"
const ACCESS_TOKEN = (typeof CHANNEL_ACCESS_TOKEN !== 'undefined') ? CHANNEL_ACCESS_TOKEN : null;

/**
 * ฟังก์ชันหลักสำหรับจัดการการบันทึกข้อมูลนัดหมาย
 */
function handleAppointmentBooking(userId, userMessage, token) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // ใช้ PropertiesService แทน global variable เพื่อเก็บข้อมูลชั่วคราว
  const userProperties = PropertiesService.getScriptProperties();
  const userDataKey = 'appointment_' + userId;
  
  // เริ่มต้นการบันทึกข้อมูลนัดหมาย - เปลี่ยนจาก "บันทึกข้อมูล" เป็น "ลงเวลานัดหมาย"
  if (userMessage.toLowerCase() === "ลงเวลานัดหมาย") {
    userProperties.setProperty(userDataKey, JSON.stringify({ step: 'waiting_date' }));
    return replyMessage(token, "📅 กรุณาระบุวันที่นัดหมาย\n\nรูปแบบที่รองรับ:\n• dd/mm/yyyy (เช่น 15/06/2568)\n• dd.mm.yyyy (เช่น 15.06.2568)\n• dd mm yyyy (เช่น 15 06 2568)\n• dd ม.ค. yyyy (เช่น 15 ม.ค. 2568)\n• dd มกราคม yyyy (เช่น 15 มกราคม 2568)\n\n💡 หมายเหตุ: ปีให้ใช้แบบ พ.ศ. (เช่น 2568)\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");
  }

  // ตรวจสอบว่าผู้ใช้อยู่ในขั้นตอนการบันทึกข้อมูลหรือไม่
  const userDataString = userProperties.getProperty(userDataKey);
  if (!userDataString) {
    return null; // ไม่ใช่การบันทึกข้อมูลนัดหมาย
  }

  let userData;
  try {
    userData = JSON.parse(userDataString);
  } catch (error) {
    Logger.log('Error parsing user data: ' + error);
    userProperties.deleteProperty(userDataKey);
    return replyMessage(token, "❌ เกิดข้อผิดพลาดในระบบ กรุณาเริ่มใหม่โดยพิมพ์ 'ลงเวลานัดหมาย'");
  }

  Logger.log('Current step: ' + userData.step + ', User message: ' + userMessage);

  // ตรวจสอบคำสั่งยกเลิก
  if (userMessage.toLowerCase() === "ยกเลิก") {
    userProperties.deleteProperty(userDataKey);
    return replyMessage(token, "❌ ยกเลิกการลงเวลานัดหมายเรียบร้อยแล้ว\n\nสามารถพิมพ์ข้อความได้ปกติ หรือพิมพ์ 'ลงเวลานัดหมาย' เพื่อเริ่มใหม่");
  }

  switch (userData.step) {
    case 'waiting_date':
      const parsedDate = parseThaiDate(userMessage);
      Logger.log('Parsed date: ' + parsedDate);
      if (parsedDate) {
        userData.date = parsedDate;
        userData.step = 'waiting_time';
        userProperties.setProperty(userDataKey, JSON.stringify(userData));
        return replyMessage(token, "⏰ กรุณาระบุเวลานัดหมาย\n\nรูปแบบที่รองรับ:\n• hh:mm (เช่น 09:30)\n• hh.mm (เช่น 09.30)\n• h:mm (เช่น 9:30)\n• h.mm (เช่น 9.30)\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");
      } else {
        return replyMessage(token, "❌ รูปแบบวันที่ไม่ถูกต้อง กรุณาลองใหม่\n\nตัวอย่างที่ถูกต้อง:\n• 15/06/2568\n• 15 มิถุนายน 2568\n• 15 มิ.ย. 2568\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");
      }

    case 'waiting_time':
      const parsedTime = parseTime(userMessage);
      Logger.log('Parsed time: ' + parsedTime);
      if (parsedTime) {
        userData.time = parsedTime;
        userData.step = 'waiting_hn';
        userProperties.setProperty(userDataKey, JSON.stringify(userData));
        return replyMessage(token, "🏥 กรุณาระบุเลข HN ของผู้ป่วย\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");
      } else {
        return replyMessage(token, "❌ รูปแบบเวลาไม่ถูกต้อง กรุณาลองใหม่\n\nตัวอย่าง:\n• 09:30\n• 9.30\n• 14:00\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");
      }

    case 'waiting_hn':
      userData.hn = userMessage.trim();
      userData.step = 'waiting_name';
      userProperties.setProperty(userDataKey, JSON.stringify(userData));
      return replyMessage(token, "👤 กรุณาระบุ ชื่อ-สกุล ของผู้ป่วย\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");

    case 'waiting_name':
      userData.name = userMessage.trim();
      userData.step = 'waiting_phone';
      userProperties.setProperty(userDataKey, JSON.stringify(userData));
      return replyMessage(token, "📞 กรุณาระบุเบอร์โทรศัพท์\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");

    case 'waiting_phone':
      userData.phone = userMessage.trim();
      userData.step = 'waiting_detail';
      userProperties.setProperty(userDataKey, JSON.stringify(userData));
      return replyMessage(token, "📝 กรุณาระบุรายละเอียดการนัดหมาย\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");

    case 'waiting_detail':
      userData.detail = userMessage.trim();
      userData.step = 'ready_to_save';
      userProperties.setProperty(userDataKey, JSON.stringify(userData));
      
      // แสดงข้อมูลที่กรอกทั้งหมดและแนะนำการค้นหา USER ID
      const summaryMessage = `✅ ข้อมูลนัดหมายครบถ้วนแล้ว!

📅 วันที่: ${userData.date}
⏰ เวลา: ${userData.time}
🏥 HN: ${userData.hn}
👤 ชื่อ-สกุล: ${userData.name}
📞 เบอร์โทร: ${userData.phone}
📝 รายละเอียด: ${userData.detail}

🔍 กรุณาเลือกตัวเลือกการแจ้งเตือน:

1️⃣ ส่งไลน์แจ้งเตือน: พิมพ์ "ID" ตามด้วย ชื่อลูกค้า
   ตัวอย่าง: ID กมลภพ, ID สมชาย

2️⃣ ไม่ส่งไลน์: พิมพ์ "ข้าม" เพื่อบันทึกข้อมูลทันที

❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย`;

      return replyMessage(token, summaryMessage);

    case 'ready_to_save':
      // ตรวจสอบว่าผู้ใช้พิมพ์ "ข้าม" หรือไม่
      if (userMessage.toLowerCase() === "ข้าม") {
        // บันทึกข้อมูลโดยไม่ส่งไลน์แจ้งเตือน
        const saveResult = saveAppointmentData(userData, null);
        
        // ลบข้อมูลชั่วคราว
        userProperties.deleteProperty(userDataKey);
        
        if (saveResult.success) {
          const confirmMessage = createConfirmationMessage(userData, saveResult, false);
          return replyMessage(token, confirmMessage);
        } else {
          return replyMessage(token, "❌ เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + saveResult.error);
        }
      }

      // ตรวจสอบการค้นหา USER ID (จะถูกจัดการโดยระบบค้นหา USER ID ใน DoPost.gs)
      // เมื่อได้ USER ID แล้ว ระบบจะเรียก handleFoundUserId() อัตโนมัติ
      
      return replyMessage(token, "⚠️ กรุณาเลือกตัวเลือก:\n\n1️⃣ ส่งไลน์แจ้งเตือน: ID ชื่อลูกค้า\n2️⃣ ไม่ส่งไลน์: ข้าม\n\n❌ พิมพ์ 'ยกเลิก' เพื่อยกเลิกการลงเวลานัดหมาย");

    default:
      userProperties.deleteProperty(userDataKey);
      return replyMessage(token, "❌ เกิดข้อผิดพลาดในระบบ กรุณาเริ่มใหม่โดยพิมพ์ 'ลงเวลานัดหมาย'");
  }
}

/**
 * ฟังก์ชันที่จะถูกเรียกจากระบบค้นหา USER ID เมื่อพบ USER ID แล้ว
 * ใช้สำหรับบันทึกข้อมูลอัตโนมัติ
 */
function handleFoundUserId(userId, customerUserId, token) {
  try {
    Logger.log('handleFoundUserId called - userId: ' + userId + ', customerUserId: ' + customerUserId);
    
    const userProperties = PropertiesService.getScriptProperties();
    const userDataKey = 'appointment_' + userId;
    const userDataString = userProperties.getProperty(userDataKey);
    
    if (!userDataString) {
      Logger.log('No appointment data found for user: ' + userId);
      return replyMessage(token, "❌ ไม่พบข้อมูลการนัดหมาย กรุณาเริ่มใหม่โดยพิมพ์ 'ลงเวลานัดหมาย'");
    }
    
    let userData;
    try {
      userData = JSON.parse(userDataString);
    } catch (parseError) {
      Logger.log('Error parsing appointment data: ' + parseError);
      userProperties.deleteProperty(userDataKey);
      return replyMessage(token, "❌ เกิดข้อผิดพลาดในข้อมูล กรุณาเริ่มใหม่โดยพิมพ์ 'ลงเวลานัดหมาย'");
    }
    
    // ตรวจสอบว่าอยู่ในขั้นตอน ready_to_save หรือไม่
    if (userData.step !== 'ready_to_save') {
      Logger.log('User not in ready_to_save step: ' + userData.step);
      return null; // ไม่ใช่เวลาที่เหมาะสม
    }
    
    Logger.log('Auto-saving appointment with customer ID: ' + customerUserId);
    
    // บันทึกข้อมูลลง Google Sheet และ Calendar
    const saveResult = saveAppointmentData(userData, customerUserId);
    
    // ลบข้อมูลชั่วคราว
    userProperties.deleteProperty(userDataKey);
    
    if (saveResult.success) {
      const confirmMessage = createConfirmationMessage(userData, saveResult, true);
      
      // ส่งไลน์แจ้งเตือนไปยังลูกค้า
      const notificationResult = sendAppointmentNotificationToCustomer(customerUserId, userData);
      
      return replyMessage(token, confirmMessage);
    } else {
      return replyMessage(token, "❌ เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + saveResult.error);
    }
    
  } catch (error) {
    Logger.log('Error in handleFoundUserId: ' + error.toString());
    return replyMessage(token, "❌ เกิดข้อผิดพลาดในระบบ: " + error.toString());
  }
}

/**
 * ฟังก์ชันดึง USER ID ของลูกค้าจากคอลัมน์ G
 */
function getCustomerUserIdFromColumnG(requesterId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userIdSheet = ss.getSheetByName("USER ID");
    const lastRow = userIdSheet.getLastRow();
    
    // หาแถวของผู้ที่ขอข้อมูล
    for (let i = 2; i <= lastRow; i++) {
      const userId = userIdSheet.getRange(i, 1).getValue();
      if (userId && userId.toString() === requesterId.toString()) {
        const customerUserId = userIdSheet.getRange(i, 7).getValue(); // คอลัมน์ G
        Logger.log('Found customer USER ID in column G: ' + customerUserId);
        return customerUserId ? customerUserId.toString() : null;
      }
    }
    
    Logger.log('No customer USER ID found in column G for requester: ' + requesterId);
    return null;
    
  } catch (error) {
    Logger.log('Error getting customer USER ID from column G: ' + error.toString());
    return null;
  }
}

/**
 * ฟังก์ชันส่งไลน์แจ้งเตือนไปยังลูกค้า
 */
function sendAppointmentNotificationToCustomer(customerUserId, userData) {
  try {
    Logger.log('Sending notification to customer: ' + customerUserId);
    
    // ตรวจสอบ ACCESS_TOKEN
    if (!ACCESS_TOKEN) {
      Logger.log('ACCESS_TOKEN not available');
      return { success: false, error: 'ACCESS_TOKEN not available' };
    }
    
    const notificationMessage = `🏥 แจ้งเตือนการนัดหมาย

สวัสดีค่ะ คุณมีการนัดหมายดังนี้:

📅 วันที่: ${userData.date}
⏰ เวลา: ${userData.time}
🏥 HN: ${userData.hn}
👤 ชื่อ-สกุล: ${userData.name}
📝 รายละเอียด: ${userData.detail}

กรุณามาตามเวลาที่นัดหมายค่ะ
หากมีข้อสงสัยติดต่อ: ${userData.phone}

ขอบคุณค่ะ 🙏`;

    // ส่งข้อความไปยังลูกค้า
    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + ACCESS_TOKEN
      },
      method: 'POST',
      payload: JSON.stringify({
        to: customerUserId,
        messages: [{
          type: "text",
          text: notificationMessage
        }]
      })
    });

    Logger.log('Notification sent to customer, response code: ' + response.getResponseCode());
    
    return {
      success: true,
      responseCode: response.getResponseCode()
    };
    
  } catch (error) {
    Logger.log('Error sending notification to customer: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ฟังก์ชันแปลงวันที่ภาษาไทยเป็นรูปแบบมาตรฐาน
 */
function parseThaiDate(dateString) {
  const cleanDate = dateString.trim();
  Logger.log('Input date string: ' + cleanDate);
  
  // รูปแบบ dd/mm/yyyy หรือ dd.mm.yyyy หรือ dd mm yyyy
  const numericPattern = /^(\d{1,2})[\s\/\.](\d{1,2})[\s\/\.](\d{4})$/;
  const numericMatch = cleanDate.match(numericPattern);
  
  if (numericMatch) {
    const day = parseInt(numericMatch[1]);
    const month = parseInt(numericMatch[2]);
    const year = parseInt(numericMatch[3]);
    
    Logger.log('Numeric match - Day: ' + day + ', Month: ' + month + ', Year: ' + year);
    
    // แปลง พ.ศ. เป็น ค.ศ.
    const gregorianYear = year > 2400 ? year - 543 : year;
    
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      const result = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
      Logger.log('Returning numeric result: ' + result);
      return result;
    }
  }
  
  // รูปแบบ dd เดือน yyyy (เช่น 15 มกราคม 2568 หรือ 15 ม.ค. 2568)
  const thaiMonthPattern = /^(\d{1,2})\s+(.+?)\s+(\d{4})$/;
  const thaiMatch = cleanDate.match(thaiMonthPattern);
  
  if (thaiMatch) {
    const day = parseInt(thaiMatch[1]);
    const monthName = thaiMatch[2].trim();
    const year = parseInt(thaiMatch[3]);
    
    Logger.log('Thai match - Day: ' + day + ', Month name: ' + monthName + ', Year: ' + year);
    
    // หาเดือนจากชื่อเดือนไทย
    const monthIndex = findThaiMonth(monthName);
    Logger.log('Month index found: ' + monthIndex);
    
    if (monthIndex !== -1 && day >= 1 && day <= 31) {
      const result = `${String(day).padStart(2, '0')}/${String(monthIndex + 1).padStart(2, '0')}/${year}`;
      Logger.log('Returning Thai result: ' + result);
      return result;
    }
  }
  
  Logger.log('No valid date format found');
  return null;
}

/**
 * ฟังก์ชันค้นหาเดือนภาษาไทย
 */
function findThaiMonth(monthName) {
  // ใช้ MONTH_NAMES_TH จาก Constants หรือ Local array
  const fullMonths = (typeof MONTH_NAMES_TH !== 'undefined') ? MONTH_NAMES_TH : MONTH_NAMES_TH_LOCAL;
  const shortMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 
                     'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  
  Logger.log('Looking for month: ' + monthName);
  
  // ค้นหาในชื่อเต็ม
  let index = fullMonths.findIndex(month => month === monthName);
  if (index !== -1) {
    Logger.log('Found full month at index: ' + index);
    return index;
  }
  
  // ค้นหาในชื่อย่อ
  index = shortMonths.findIndex(month => month === monthName);
  if (index !== -1) {
    Logger.log('Found short month at index: ' + index);
    return index;
  }
  
  // ค้นหาแบบบางส่วน (เผื่อพิมพ์ผิด)
  for (let i = 0; i < fullMonths.length; i++) {
    if (fullMonths[i].includes(monthName) || monthName.includes(fullMonths[i])) {
      Logger.log('Found partial match at index: ' + i);
      return i;
    }
  }
  
  Logger.log('No month found');
  return -1;
}

/**
 * ฟังก์ชันแปลงเวลา
 */
function parseTime(timeString) {
  const cleanTime = timeString.trim();
  
  // รูปแบบ hh:mm หรือ h:mm หรือ hh.mm หรือ h.mm
  const timePattern = /^(\d{1,2})[\:\.](\d{2})$/;
  const match = cleanTime.match(timePattern);
  
  if (match) {
    const hour = parseInt(match[1]);
    const minute = parseInt(match[2]);
    
    if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
      return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:00`;
    }
  }
  
  return null;
}

/**
 * ฟังก์ชันบันทึกข้อมูลลง Google Sheet และ Calendar
 */
function saveAppointmentData(userData, customerUserId) {
  try {
    Logger.log('=== Starting saveAppointmentData ===');
    Logger.log('User data: ' + JSON.stringify(userData));
    Logger.log('Customer User ID: ' + customerUserId);
    
    // ตรวจสอบ SHEET_ID ก่อน
    if (typeof SHEET_ID === 'undefined') {
      throw new Error('SHEET_ID is not defined');
    }
    Logger.log('SHEET_ID: ' + SHEET_ID);
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Spreadsheet opened successfully');
    
    // หาเดือนจากวันที่ที่บันทึก
    const dateParts = userData.date.split('/');
    const month = parseInt(dateParts[1]);
    Logger.log('Month number: ' + month);
    
    // ใช้ MONTH_NAMES_TH จาก Constants หรือ Local array
    const monthNames = (typeof MONTH_NAMES_TH !== 'undefined') ? MONTH_NAMES_TH : MONTH_NAMES_TH_LOCAL;
    const monthName = monthNames[month - 1];
    Logger.log('Month name: ' + monthName);
    
    if (!monthName) {
      throw new Error('ไม่สามารถหาชื่อเดือนได้ Month index: ' + (month - 1));
    }
    
    // หาหรือสร้าง Sheet ของเดือนนั้น
    let monthSheet = ss.getSheetByName(monthName);
    Logger.log('Existing month sheet found: ' + (monthSheet !== null));
    
    if (!monthSheet) {
      Logger.log('Creating new sheet for month: ' + monthName);
      monthSheet = ss.insertSheet(monthName);
      
      // สร้างหัวตาราง
      const headers = [['วันที่', 'เวลา', 'HN', 'ชื่อ-สกุล', 'รายละเอียด', 'เบอร์โทร', 'Customer ID']];
      monthSheet.getRange(1, 1, 1, 7).setValues(headers);
      Logger.log('Header row created');
    }
    
    // หาแถวว่างถัดไป
    const lastRow = monthSheet.getLastRow();
    const newRow = lastRow + 1;
    Logger.log('Writing to row: ' + newRow);
    
    // เตรียมข้อมูลสำหรับบันทึก (บันทึกทีเดียวทั้งแถว)
    const rowData = [
      userData.date,        // A: วันที่
      userData.time,        // B: เวลา  
      userData.hn,          // C: HN
      userData.name,        // D: ชื่อ-สกุล
      userData.detail,      // E: รายละเอียด
      userData.phone,       // F: เบอร์โทร
      customerUserId        // G: Customer ID
    ];
    
    Logger.log('Row data prepared: ' + JSON.stringify(rowData));
    
    // บันทึกข้อมูลทั้งแถวในครั้งเดียว
    try {
      monthSheet.getRange(newRow, 1, 1, 7).setValues([rowData]);
      Logger.log('✅ All data written to sheet successfully');
      
      // ตรวจสอบว่าข้อมูลถูกบันทึกจริงหรือไม่
      const verifyData = monthSheet.getRange(newRow, 1, 1, 7).getValues()[0];
      Logger.log('Verification - data in sheet: ' + JSON.stringify(verifyData));
      
    } catch (writeError) {
      Logger.log('❌ Error writing to sheet: ' + writeError.toString());
      Logger.log('Write error stack: ' + writeError.stack);
      throw writeError;
    }
    
    // บันทึกลง Google Calendar
    Logger.log('Starting calendar save...');
    const calendarResult = saveToGoogleCalendar(userData);
    Logger.log('Calendar result: ' + JSON.stringify(calendarResult));
    
    const result = {
      success: true,
      sheetRow: newRow,
      calendarEvent: calendarResult,
      customerUserId: customerUserId,
      monthSheet: monthName,
      sheetData: rowData
    };
    
    Logger.log('=== saveAppointmentData completed successfully ===');
    Logger.log('Final result: ' + JSON.stringify(result));
    
    return result;
    
  } catch (error) {
    Logger.log('❌ Error saving appointment data: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return {
      success: false,
      error: error.toString(),
      errorStack: error.stack
    };
  }
}

/**
 * ฟังก์ชันบันทึกลง Google Calendar
 */
function saveToGoogleCalendar(userData) {
  try {
    // แปลงวันที่และเวลา
    const dateParts = userData.date.split('/');
    const day = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]);
    const year = parseInt(dateParts[2]);
    
    // แปลง พ.ศ. เป็น ค.ศ. สำหรับ Calendar
    const gregorianYear = year > 2400 ? year - 543 : year;
    
    const timeParts = userData.time.split(':');
    const hour = parseInt(timeParts[0]);
    const minute = parseInt(timeParts[1]);
    
    // สร้างวันที่และเวลาเริ่มต้น
    const startDate = new Date(gregorianYear, month - 1, day, hour, minute, 0);
    
    // สร้างเวลาสิ้นสุด (เพิ่ม 3 ชั่วโมง)
    const endDate = new Date(startDate.getTime() + (3 * 60 * 60 * 1000));
    
    // สร้างชื่อกิจกรรม
    const eventTitle = `นัดหมาย: ${userData.name}`;
    
    // สร้างรายละเอียด
    const description = `รหัส HN: ${userData.hn}
ชื่อ-สกุล: ${userData.name}
มีนัดหมอ เพื่อ: ${userData.detail}
เบอร์โทร: ${userData.phone}`;
    
    // บันทึกลง Calendar
    const calendar = CalendarApp.getCalendarById(APPOINTMENT_CALENDAR_ID);
    const event = calendar.createEvent(eventTitle, startDate, endDate, {
      description: description,
      location: 'โรงพยาบาล'
    });
    
    Logger.log('Calendar event created: ' + event.getId());
    
    return {
      success: true,
      eventId: event.getId(),
      startTime: startDate,
      endTime: endDate
    };
    
  } catch (error) {
    Logger.log('Error creating calendar event: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ฟังก์ชันสร้างข้อความยืนยันการบันทึก
 */
function createConfirmationMessage(userData, saveResult, withNotification = true) {
  let message = `✅ ได้บันทึกข้อมูลเรียบร้อยแล้ว!

📅 วันที่: ${userData.date}
⏰ เวลา: ${userData.time}
🏥 HN: ${userData.hn}
👤 ชื่อ-สกุล: ${userData.name}
📞 เบอร์โทร: ${userData.phone}
📝 รายละเอียด: ${userData.detail}

ข้อมูลได้ถูกบันทึกลงใน Google Sheet และ Google Calendar เรียบร้อยแล้ว`;

  if (withNotification && saveResult.customerUserId) {
    message += `\n\n📱 ส่งไลน์แจ้งเตือนไปยังลูกค้าเรียบร้อยแล้ว`;
  } else if (!withNotification) {
    message += `\n\n📝 บันทึกข้อมูลโดยไม่ส่งไลน์แจ้งเตือน`;
  }

  return message;
}

/**
 * ฟังก์ชันทดสอบระบบ (สำหรับ Admin)
 */
function testAppointmentSystem() {
  const testData = {
    date: "15/06/2568",
    time: "09:30:00",
    hn: "HN123456",
    name: "ทดสอบ ระบบ",
    phone: "081-234-5678",
    detail: "ตรวจสุขภาพประจำปี"
  };
  
  const result = saveAppointmentData(testData, "test_customer_id");
  Logger.log('Test result: ' + JSON.stringify(result));
  
  return result;
}

/**
 * ฟังก์ชันทดสอบเฉพาะการบันทึก Google Sheet
 */
function testSheetWriteOnly() {
  try {
    Logger.log('=== Testing Sheet Write Only ===');
    
    // ตรวจสอบ SHEET_ID
    if (typeof SHEET_ID === 'undefined') {
      Logger.log('❌ SHEET_ID is not defined');
      return { success: false, error: 'SHEET_ID not defined' };
    }
    
    Logger.log('✅ SHEET_ID found: ' + SHEET_ID);
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('✅ Spreadsheet opened');
    
    // ลองสร้าง sheet ทดสอบ
    const testSheetName = 'ทดสอบ_' + new Date().getTime();
    const testSheet = ss.insertSheet(testSheetName);
    Logger.log('✅ Test sheet created: ' + testSheetName);
    
    // ลองเขียนข้อมูลทดสอบ
    const testData = [['A1', 'B1', 'C1'], ['A2', 'B2', 'C2']];
    testSheet.getRange(1, 1, 2, 3).setValues(testData);
    Logger.log('✅ Test data written');
    
    // ลองอ่านข้อมูลกลับมา
    const readData = testSheet.getRange(1, 1, 2, 3).getValues();
    Logger.log('✅ Test data read back: ' + JSON.stringify(readData));
    
    // ลบ sheet ทดสอบ
    ss.deleteSheet(testSheet);
    Logger.log('✅ Test sheet deleted');
    
    return { 
      success: true, 
      message: 'Sheet write test passed',
      testData: readData 
    };
    
  } catch (error) {
    Logger.log('❌ Sheet test error: ' + error.toString());
    return { 
      success: false, 
      error: error.toString(),
      stack: error.stack 
    };
  }
}

/**
 * ฟังก์ชันตรวจสอบ Constants ทั้งหมด
 */
function checkConstants() {
  const constants = {
    SHEET_ID: typeof SHEET_ID !== 'undefined' ? SHEET_ID : 'NOT DEFINED',
    ACCESS_TOKEN: typeof ACCESS_TOKEN !== 'undefined' ? 'DEFINED' : 'NOT DEFINED',
    CHANNEL_ACCESS_TOKEN: typeof CHANNEL_ACCESS_TOKEN !== 'undefined' ? 'DEFINED' : 'NOT DEFINED',
    APPOINTMENT_CALENDAR_ID: typeof APPOINTMENT_CALENDAR_ID !== 'undefined' ? APPOINTMENT_CALENDAR_ID : 'NOT DEFINED',
    MONTH_NAMES_TH: typeof MONTH_NAMES_TH !== 'undefined' ? 'DEFINED' : 'NOT DEFINED'
  };
  
  Logger.log('Constants check: ' + JSON.stringify(constants));
  return constants;
}

/**
 * ฟังก์ชันล้างข้อมูลการนัดหมายที่ค้างอยู่ (สำหรับทดสอบ)
 */
function clearAllAppointmentData() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    let clearedCount = 0;
    for (const key in allProperties) {
      if (key.startsWith('appointment_')) {
        properties.deleteProperty(key);
        clearedCount++;
      }
    }
    
    Logger.log('Cleared ' + clearedCount + ' appointment properties');
    return { success: true, clearedCount: clearedCount };
    
  } catch (error) {
    Logger.log('Error clearing appointment data: ' + error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ฟังก์ชันดูข้อมูลการนัดหมายที่ค้างอยู่ (สำหรับดีบัก)
 */
function viewPendingAppointments() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    const appointments = {};
    for (const key in allProperties) {
      if (key.startsWith('appointment_')) {
        try {
          appointments[key] = JSON.parse(allProperties[key]);
        } catch (parseError) {
          appointments[key] = allProperties[key];
        }
      }
    }
    
    Logger.log('Pending appointments: ' + JSON.stringify(appointments));
    return appointments;
    
  } catch (error) {
    Logger.log('Error viewing pending appointments: ' + error);
    return { error: error.toString() };
  }
}

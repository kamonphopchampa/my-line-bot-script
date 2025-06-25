/**
 * ส่งไลน์.gs - ฟังก์ชันสนับสนุนสำหรับระบบ LINE Bot
 * (ลบ doPost ออกแล้ว เพราะย้ายไปไฟล์ "รวม DoPost.gs")
 */

// LINE Messaging API Channel Access Token
const CHANNEL_ACCESS_TOKEN = 'xxxxxxxxxxx;

// Google Sheet ID
const SHEET_ID = 'xxxxxxxxxxxx';

// Google Calendar ID สำหรับบันทึกนัดหมาย
const APPOINTMENT_CALENDAR_ID = 'xxxxxxxxxxxxxxx';

// Month names in Thai
const MONTH_NAMES_TH = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// Day names in Thai
const DAY_NAMES_TH = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

// *** ลบฟังก์ชัน doPost ออกแล้ว เพราะย้ายไปไฟล์ "รวม DoPost.gs" ***

/**
 * Helper function to format date in Thai format
 */
function formatThaiDate(date) {
  const day = date.getDate();
  const month = MONTH_NAMES_TH[date.getMonth()];
  // Convert to Buddhist Era (BE) by adding 543 to the year
  const year = date.getFullYear() + 543;
  const dayName = DAY_NAMES_TH[date.getDay()];
  
  return `วัน${dayName}ที่ ${day} ${month} ${year} `;
}

/**
 * Helper function to pad single digits with zero
 */
function padZero(num) {
  return num < 10 ? '0' + num : num;
}

/**
 * Function to get user profile information
 */
function getUserProfiles(userId) {
  const url = "https://api.line.me/v2/bot/profile/" + userId;
  const lineHeader = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN
  };

  const options = {
    "method": "GET",
    "headers": lineHeader
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseJson = JSON.parse(response);
  const displayName = responseJson.displayName;
  const pictureUrl = responseJson.pictureUrl || "";

  return [displayName, pictureUrl];
}

/**
 * Function to get keyword responses from the keyword sheet
 */
function getKeywordResponse(userMessage, keywordSheet) {
  const lastRow = keywordSheet.getLastRow();
  
  // Check column A & B for questions and answers
  for (let i = 1; i <= lastRow; i++) {
    const question1 = keywordSheet.getRange(i, 1).getValue();
    const answer1 = keywordSheet.getRange(i, 2).getValue();
    if (userMessage.toLowerCase() === question1.toLowerCase()) {
      return answer1;
    }
  }

  // Check column E & F for questions and answers
  for (let i = 1; i <= lastRow; i++) {
    const question2 = keywordSheet.getRange(i, 5).getValue();
    const answer2 = keywordSheet.getRange(i, 6).getValue();
    if (userMessage.toLowerCase() === question2.toLowerCase()) {
      return answer2;
    }
  }

  return null;
}

/**
 * Function to reply with text message
 */
function replyMessage(token, replyText) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const lineHeader = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN
  };

  const postData = {
    "replyToken": token,
    "messages": [{
      "type": "text",
      "text": replyText
    }]
  };

  const options = {
    "method": "POST",
    "headers": lineHeader,
    "payload": JSON.stringify(postData)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log("Status code: " + response.getResponseCode());
    if (response.getResponseCode() === 200) {
      Logger.log("Sending message completed.");
    }
    return ContentService.createTextOutput(JSON.stringify({'status': 'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log(error.name + "：" + error.message);
    return ContentService.createTextOutput(JSON.stringify({'status': 'error', 'message': error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Function to create a monthly report Flex Message
 */
function createMonthlyReport(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return null;
  }
  
  // Get report title from B1:D1
  const reportTitle = sheet.getRange('B1').getValues()[0].join(' ').trim();
  
  // === รายรับ ===
  // Get income data from A2:D15
  const incomeData = sheet.getRange('A2:B121').getValues();
  
  // Get income summary data from B16:D20
  const incomeSummary = sheet.getRange('A123:B123').getValues();
  
  // === รายจ่าย ===
  // Get expense data from A22:D31
  const expenseData = sheet.getRange('A125:B144').getValues();
  
  // Get expense summary data from B32:D40
  const expenseSummary = sheet.getRange('A147:B149').getValues();
  
  // Format current date in Thai format
  const now = new Date();
  const thaiDate = formatThaiDate(now);
  
  // Filter out empty rows
  const filteredIncomeData = incomeData.filter(row => row.some(cell => cell !== ''));
  const filteredIncomeSummary = incomeSummary.filter(row => row.some(cell => cell !== ''));
  const filteredExpenseData = expenseData.filter(row => row.some(cell => cell !== ''));
  const filteredExpenseSummary = expenseSummary.filter(row => row.some(cell => cell !== ''));
  
  // Create content for income report section - limit to max 20 rows
  const maxIncomeRows = Math.min(filteredIncomeData.length, 20);
  const incomeContents = [];
  
  for (let i = 0; i < maxIncomeRows; i++) {
    const row = filteredIncomeData[i];
    // แก้ไขเพื่อให้แสดงข้อความเต็มโดยการลดขนาดตัวอักษรและเพิ่ม wrap: true
    incomeContents.push({
      "type": "box",
      "layout": "vertical",
      "contents": [
        {
          "type": "box",
          "layout": "horizontal",
          "contents": [
            {
              "type": "text",
              "text": String(row[0] || " "),
              "size": "xs",
              "color": "#555555",
              "flex": 0,
              "wrap": true
            },
            {
              "type": "text",
              "text": String(row[1] || " "),
              "size": "xs",
              "color": "#111111",
              "flex": 3,
              "wrap": true
            }
          ]
        }
      ],
      "margin": "sm"
    });
  }

  // Create content for income summary section
  const incomeSummaryContents = [];
  const maxIncomeSummaryRows = Math.min(filteredIncomeSummary.length, 3);
  
  for (let i = 0; i < maxIncomeSummaryRows; i++) {
    const row = filteredIncomeSummary[i];
    incomeSummaryContents.push({
      "type": "box",
      "layout": "horizontal",
      "contents": [
        {
          "type": "text",
          "text": String(row[0] || "-"),
          "size": "sm",
          "color": "#555555",
          "flex": 0
        },
        {
          "type": "text",
          "text": String(row[1] || "-"),
          "size": "sm",
          "color": "#111111",
          "flex": 3
        }
      ]
    });
  }
  
  // Create content for expense report section - limit to max 20 rows
  const maxExpenseRows = Math.min(filteredExpenseData.length, 20);
  const expenseContents = [];
  
  for (let i = 0; i < maxExpenseRows; i++) {
    const row = filteredExpenseData[i];
    expenseContents.push({
      "type": "box",
      "layout": "horizontal",
      "contents": [
        {
          "type": "text",
          "text": String(row[0] || "-"),
          "size": "sm",
          "color": "#555555",
          "flex": 0
        },
        {
          "type": "text",
          "text": String(row[1] || "-"),
          "size": "sm",
          "color": "#111111",
          "flex": 3
        }
      ]
    });
  }
  
  // Create content for expense summary section
  const expenseSummaryContents = [];
  const maxExpenseSummaryRows = Math.min(filteredExpenseSummary.length, 3);
  
  for (let i = 0; i < maxExpenseSummaryRows; i++) {
    const row = filteredExpenseSummary[i];
    expenseSummaryContents.push({
      "type": "box",
      "layout": "horizontal",
      "contents": [
        {
          "type": "text",
          "text": String(row[0] || "-"),
          "size": "sm",
          "color": "#555555",
          "flex": 0
        },
        {
          "type": "text",
          "text": String(row[1] || "-"),
          "size": "sm",
          "color": "#111111",
          "flex": 3
        }
      ]
    });
  }
  
  // Construct the Flex Message
  const flexMessage = {
    "type": "flex",
    "altText": "รายงาน: " + reportTitle,
    "contents": {
      "type": "bubble",
      "size": "mega",
      "header": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": reportTitle,
            "weight": "bold",
            "size": "md",
            "color": "#efe606",
            "wrap": true,
            "maxLines": 3
          },
          {
            "type": "text",
            "text": thaiDate,
            "size": "sm",
            "color": "#efe606",
            "margin": "md",
            "wrap": true
          }
        ],
        "backgroundColor": "#0367D3",
        "paddingTop": "19px",
        "paddingAll": "12px",
        "paddingBottom": "16px"
      },
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "📈 นัดหมายพรุ่งนี้",
            "weight": "bold",
            "size": "md",
            "margin": "md",
            "color": "#0367D3"
          },
          {
            "type": "separator",
            "margin": "sm"
          },
          {
            "type": "box",
            "layout": "vertical",
            "margin": "lg",
            "spacing": "sm",
            "contents": incomeContents.length > 0 ? incomeContents : [{
              "type": "text",
              "text": " ",
              "size": "sm",
              "color": "#999999",
              "align": "center"
            }]
          },
          {
            "type": "text",
            "text": " ",
            "weight": "bold",
            "size": "md",
            "margin": "xl",
            "color": "#0367D3",
          },
          {
            "type": "box",
            "layout": "vertical",
            "margin": "lg",
            "spacing": "sm",
            "contents": incomeSummaryContents.length > 0 ? incomeSummaryContents : [{
              "type": "text",
              "text": " ",
              "size": "sm",
              "color": "#999999",
              "align": "center"
            }]
          },
          {
            "type": "separator",
            "margin": "xl"
          },
          {
            "type": "text",
            "text": " ",
            "weight": "bold",
            "size": "md",
            "margin": "xl",
            "color": "#D30347"
          },
          {
            "type": "separator",
            "margin": "sm"
          },
          {
            "type": "box",
            "layout": "vertical",
            "margin": "lg",
            "spacing": "sm",
            "contents": expenseContents.length > 0 ? expenseContents : [{
              "type": "text",
              "text": " ",
              "size": "sm",
              "color": "#999999",
              "align": "center"
            }]
          },
          {
            "type": "text",
            "text": " ",
            "weight": "bold",
            "size": "md",
            "margin": "xl",
            "color": "#D30347",
          },
          {
            "type": "text",
            "text": " ",
            "weight": "bold",
            "size": "md",
            "margin": "xl",
            "color": "#0367D3",
          },
          {
            "type": "box",
            "layout": "vertical",
            "margin": "lg",
            "spacing": "sm",
            "contents": expenseSummaryContents.length > 0 ? expenseSummaryContents : [{
              "type": "text",
              "text": " ",
              "size": "sm",
              "color": "#999999",
              "align": "center"
            }]
          }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": [
          {
            "type": "box",
            "layout": "vertical",
            "cornerRadius": "md",
            "backgroundColor": "#0367D3",
            "paddingAll": "8px",
            "action": {
              "type": "uri",
              "label": "ดูข้อมูลเพิ่มเติม",
              "uri": `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=0`
            },
            "contents": [
              {
                "type": "text",
                "text": "ดูข้อมูลเพิ่มเติม",
                "color": "#efe606",
                "size": "md",
                "align": "center",
                "weight": "bold"
              }
            ]
          }
        ]
      }
    }
  };
  
  return flexMessage;
}

/**
 * Function to push a monthly report to a specific user or group
 */
function pushMonthlyReport(targetId, sheetName) {
  const flexMessage = createMonthlyReport(sheetName);
  
  if (!flexMessage) {
    Logger.log("Could not create flex message for sheet: " + sheetName);
    return false;
  }
  
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    'to': targetId,
    'messages': [flexMessage]
  };
  
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode >= 200 && responseCode < 300) {
      Logger.log('Message pushed successfully.');
      return true;
    } else {
      Logger.log('Failed to push message. Response: ' + response.getContentText());
      
      // Send simplified message if Flex message fails
      if (responseCode === 400) {
        pushSimplifiedMessage(targetId, sheetName);
      }
      
      return false;
    }
  } catch (error) {
    Logger.log('Error pushing message: ' + error);
    pushSimplifiedMessage(targetId, sheetName);
    return false;
  }
}

/**
 * Function to push a simplified message if Flex message fails
 */
function pushSimplifiedMessage(targetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    // Get report title
    const reportTitle = sheet.getRange('B1').getValues()[0].join(' ').trim();
    
    // Create simple text message
    const textMessage = {
      'type': 'text',
      'text': `📊 ${reportTitle}\n\nข้อมูลรายงานมีขนาดใหญ่เกินไปที่จะแสดงใน LINE\nกรุณาดูข้อมูลเพิ่มเติมได้ที่ลิงก์ด้านล่าง\n\nhttps://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=0`
    };
    
    const url = 'https://api.line.me/v2/bot/message/push';
    const payload = {
      'to': targetId,
      'messages': [textMessage]
    };
    
    const options = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Simplified message sent. Response: ' + response.getResponseCode());
    return true;
  } catch (error) {
    Logger.log('Error sending simplified message: ' + error);
    return false;
  }
}

/**
 * Function to create a daily trigger for sending reports
 */
function createDailyTrigger() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger to run daily at 9:00 AM
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}

/**
 * Function to send daily reports to all registered users/groups
 * (sends tomorrow's report one day in advance)
 */
function sendAdvanceDailyReport3() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const userIdSheet = ss.getSheetByName("USER ID");
  
  // วันที่ของวันพรุ่งนี้
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowDate = tomorrow.getDate(); // วันที่
  const tomorrowMonthIndex = tomorrow.getMonth(); // เดือน (0-11)
  const tomorrowMonthName = MONTH_NAMES_TH[tomorrowMonthIndex]; // ชื่อเดือนภาษาไทย
  const sheetName = "รายงาน" + tomorrowMonthName;
  
  // ตรวจสอบว่าแผ่นงานรายงานของเดือนพรุ่งนี้มีอยู่หรือไม่
  const reportSheet = ss.getSheetByName(sheetName);
  if (!reportSheet) {
    Logger.log("ไม่พบชีทรายงานของเดือน: " + tomorrowMonthName);
    return;
  }
  
  // ดึงข้อมูล ID ผู้ใช้ / กลุ่ม Line
  const lastRow = userIdSheet.getLastRow();
  
  // ดึงข้อมูลจากทั้งคอลัมน์ F และ G
  const targetIdsF = userIdSheet.getRange("F2:F" + lastRow).getValues();
  const targetIdsG = userIdSheet.getRange("G2:G" + lastRow).getValues();
  
  // สร้าง Set เพื่อเก็บ ID ที่ไม่ซ้ำกัน
  const uniqueIds = new Set();
  
  // เพิ่ม ID จากคอลัมน์ F
  for (let i = 0; i < targetIdsF.length; i++) {
    const id = targetIdsF[i][0];
    if (id && id.toString().trim() !== "") {
      uniqueIds.add(id.toString().trim());
    }
  }
  
  // เพิ่ม ID จากคอลัมน์ G
  for (let i = 0; i < targetIdsG.length; i++) {
    const id = targetIdsG[i][0];
    if (id && id.toString().trim() !== "") {
      uniqueIds.add(id.toString().trim());
    }
  }
  
  // ส่งรายงานไปยัง ID ที่ไม่ซ้ำกัน
  uniqueIds.forEach(id => {
    if (id) {
      Logger.log("กำลังส่งรายงานไปยัง ID: " + id);
      // ส่งรายงานของ "วันพรุ่งนี้"
      pushMonthlyReport(id, sheetName, tomorrowDate);
      Utilities.sleep(500); // ป้องกัน rate limit
    }
  });
  
  Logger.log("ส่งรายงานเสร็จสิ้น จำนวน ID ทั้งหมด: " + uniqueIds.size);
}

/**
 * Function to create time-based trigger for sending reports
 * Sets up daily trigger at 21:30
 */
function createTimeTrigger() {
  // Delete any existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendAdvanceDailyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new time-based trigger to run at 21:30
  ScriptApp.newTrigger('sendAdvanceDailyReport')
    .timeBased()
    .atHour(21)
    .nearMinute(30)
    .everyDays(1)
    .create();
  
  Logger.log("Daily report trigger created successfully for 21:30 every day");
}

/**
 * ฟังก์ชันเพิ่มเมนูในสเปรดชีต
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ระบบนัดหมาย')
    .addItem('sendAdvanceDailyReport', 'updateTomorrowAppointments')
    .addToUi();
}

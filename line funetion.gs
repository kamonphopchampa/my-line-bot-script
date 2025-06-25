/**
 * ฟังก์ชันช่วยเหลือ - เพิ่ม 0 หน้าเลข
 */


/**
 * ส่งข้อความตอบกลับธรรมดา
 */
function replyMessage(token, message) {
  try {
    const replyData = {
      replyToken: token,
      messages: [typeof message === 'string' ? { type: 'text', text: message } : message]
    };
    
    const response = UrlFetchApp.fetch('https://api.line.merr/v2/bot/message/reply', {
      headers: {
        'Content-Type': 'applicatdion/json',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
      },
      method: 'POST',
      payload: JSON.stringify(replyData),
      muteHttpExceptions: true
    });
    
    Logger.log("Reply Message Status: " + response.getResponseCode());
    
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      Logger.log("✅ Text message sent successfully");
      return true;
    } else {
      Logger.log("❌ Error sending text message:", response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log("Error in replyMessage: " + error.message);
    return false;
  }
}

/**
 * ส่งข้อความ Flex Message พร้อม Fallback
 * รวมฟีเจอร์จากทั้ง 2 ฟังก์ชันเดิม + ปรับปรุงเพิ่มเติม
 */
function replyFlexMessage(token, flexMessage) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const lineHeader = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN
  };
  
  const postData = {
    "replyToken": token,
    "messages": [flexMessage]
  };
  
  const options = {
    "method": "POST",
    "headers": lineHeader,
    "payload": JSON.stringify(postData),
    "muteHttpExceptions": true  // เพิ่มเพื่อจัดการ error ได้ดีกว่า
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    Logger.log("Flex Message Status code: " + responseCode);
    
    if (responseCode >= 200 && responseCode < 300) {
      Logger.log("✅ Flex message sent successfully.");
      return ContentService.createTextOutput(JSON.stringify({'status': 'ok'}))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      // Flex message ล้มเหลว - ส่งข้อความธรรมดาแทน
      Logger.log("❌ Flex message failed (Status: " + responseCode + "), sending fallback message");
      Logger.log("Response: " + response.getContentText());
      
      const fallbackSuccess = replyMessage(token, 
        "ไม่สามารถแสดงรายงานในรูปแบบ Flex ได้ กรุณาดูรายงานที่ Sheet: " + 
        "https://docs.google.com/spreadsheets/d/" + SHEET_ID);
      
      return ContentService.createTextOutput(JSON.stringify({
        'status': fallbackSuccess ? 'ok_fallback' : 'error',
        'message': 'Flex failed, fallback ' + (fallbackSuccess ? 'sent' : 'failed')
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    Logger.log("❌ Exception in replyFlexMessage: " + error.name + " - " + error.message);
    
    // พยายามส่งข้อความธรรมดาแทน
    try {
      const fallbackSuccess = replyMessage(token, 
        "เกิดข้อผิดพลาดในการส่งรายงาน กรุณาดูรายงานที่ Sheet: " + 
        "https://docs.google.com/spreadsheets/d/" + SHEET_ID);
      
      return ContentService.createTextOutput(JSON.stringify({
        'status': fallbackSuccess ? 'error_fallback_sent' : 'error',
        'message': error.message
      })).setMimeType(ContentService.MimeType.JSON);
      
    } catch (fallbackError) {
      Logger.log("❌ Fallback also failed: " + fallbackError.message);
      return ContentService.createTextOutput(JSON.stringify({
        'status': 'error', 
        'message': 'Both flex and fallback failed: ' + error.message
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
}

/**
 * ส่งข้อความ Push (ไม่ต้องมี replyToken)
 */
function pushMessage(userId, message) {
  try {
    const pushData = {
      to: userId,
      messages: [typeof message === 'string' ? { type: 'text', text: message } : message]
    };
    
    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
      },
      method: 'POST',
      payload: JSON.stringify(pushData),
      muteHttpExceptions: true
    });
    
    Logger.log("Push Message Status: " + response.getResponseCode());
    
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      Logger.log("✅ Push message sent successfully");
      return true;
    } else {
      Logger.log("❌ Error sending push message:", response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log("Error in pushMessage: " + error.message);
    return false;
  }
}

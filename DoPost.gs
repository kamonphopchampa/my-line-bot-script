/**
 * รวม DoPost.gs - ไฟล์หลักสำหรับจัดการ Webhook (เขียนใหม่)
 * หมายเหตุ: ตัวแปร Constants ทั้งหมดอยู่ในไฟล์ "ส่งไลน์.gs" แล้ว
 */

/**
 * Main function to handle POST requests from LINE - เขียนใหม่
 */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userIdSheet = ss.getSheetByName("USER ID");
    const keywordSheet = ss.getSheetByName("คีย์เวิร์ด");
    
    const requestJSON = e.postData.contents;
    const requestObj = JSON.parse(requestJSON).events[0];
    const token = requestObj.replyToken;

    // Skip if this is a different event type or doesn't have a reply token
    if (!token || token === '00000000000000000000000000000000') {
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }

    let userMessage = '';
    if (requestObj.message && requestObj.message.type === 'text') {
      userMessage = requestObj.message.text.trim();
    }

    const userId = requestObj.source.userId || requestObj.source.groupId;

    console.log('📨 Received message:', userMessage, 'from user:', userId);

    // ===== 1. ระบบบันทึกผู้ใช้ใหม่อัตโนมัติ =====
    if (requestObj.type === 'message' || requestObj.type === 'follow') {
      const currentUserId = requestObj.source.userId;
      
      if (currentUserId) {
        // ตรวจสอบว่า userId นี้มีใน Sheet แล้วหรือยัง
        const lastRow = userIdSheet.getLastRow();
        const existingUserIds = lastRow >= 2 ? 
          userIdSheet.getRange('A2:A' + lastRow).getValues().flat() : [];
        
        if (!existingUserIds.includes(currentUserId)) {
          // ดึงข้อมูลโปรไฟล์จาก LINE API
          const profile = getUserProfile(currentUserId);
          const displayName = profile?.displayName || '';
          const pictureUrl = profile?.pictureUrl || '';

          // หาตำแหน่งแถวถัดไป (เริ่มจากแถว 2)
          const nextRow = lastRow + 1;

          // ใส่ข้อมูลในคอลัมน์ A (userId), B (ชื่อ), C (ลิงก์รูป)
          userIdSheet.getRange(`A${nextRow}`).setValue(currentUserId);
          userIdSheet.getRange(`B${nextRow}`).setValue(displayName);
          userIdSheet.getRange(`C${nextRow}`).setValue(pictureUrl);
          
          console.log('✅ บันทึกผู้ใช้ใหม่:', displayName, currentUserId);
        }
      }
    }

    // ===== 2. ระบบจัดการ Group =====
    if (requestObj.source.type === "group") {
      const groupId = requestObj.source.groupId;
      
      // Check if group exists in sheet
      const lastRow = userIdSheet.getLastRow();
      let groupExists = false;
      
      for (let i = 1; i <= lastRow; i++) {
        const existingGroupId = userIdSheet.getRange(i, 1).getValue();
        if (existingGroupId === groupId) {
          groupExists = true;
          break;
        }
      }

      // If group is new, record Group ID
      if (!groupExists) {
        userIdSheet.getRange(lastRow + 1, 1).setValue(groupId);
        userIdSheet.getRange(lastRow + 1, 2).setValue("Group");
        userIdSheet.getRange(lastRow + 1, 6).setValue(groupId);
        
        return replyMessage(token, "บอทได้รับการเชิญเข้ากลุ่มแล้ว Group ID ของกลุ่มนี้คือ: " + groupId);
      }

      // Reply with Group ID if asked
      if (userMessage.toLowerCase() === "id กลุ่มนี้คืออะไร") {
        return replyMessage(token, "Group ID ของกลุ่มนี้คือ: " + groupId);
      }

    } else if (requestObj.source.type === "user") {
      const userId = requestObj.source.userId;
      
      // Handle new friend (follow event)
      if (requestObj.type === "follow") {
        const userProfiles = getUserProfiles(userId);
        const currentDate = new Date();
        const formattedDate = formatThaiDate(currentDate);
        
        // Record user data in original columns (A-F)
        const lastRow = userIdSheet.getLastRow() + 1;
        userIdSheet.getRange(lastRow, 1).setValue(userId);
        userIdSheet.getRange(lastRow, 2).setValue(userProfiles[0]);
        userIdSheet.getRange(lastRow, 3).setValue(userProfiles[1]);
        userIdSheet.getRange(lastRow, 4).setFormula('=IMAGE(C' + lastRow + ')');
        userIdSheet.getRange(lastRow, 5).setValue(formattedDate);
        
        return replyMessage(token, "ขอบคุณที่เพิ่มเพื่อน! ท่านสามารถตรวจสอบ USER ID ได้ด้วยการพิมพ์คำว่า\n ID ของฉันคืออะไร\n ได้เลย");
      }
      
      // Reply with User ID if asked
      if (userMessage.toLowerCase() === "id ของฉันคืออะไร") {
        return replyMessage(token, "User ID ของคุณคือ: " + userId);
      }
    }

    // ===== 3. ระบบค้นหา USER ID (ความสำคัญสูงสุด) =====
    if (userMessage && userMessage.toLowerCase().startsWith("id ")) {
      const searchName = userMessage.substring(3).trim();
      console.log('🔍 Processing USER ID search for:', searchName);
      return processUserIdSearch(token, searchName, userId, userIdSheet);
    }

    // ===== 4. จัดการการเลือก USER ID ด้วยหมายเลข =====
    if (userMessage && /^[1-9]$/.test(userMessage.trim())) {
      const selectedNumber = parseInt(userMessage.trim());
      console.log('🔢 Processing number selection:', selectedNumber);
      
      const selectionResult = processNumberSelection(token, selectedNumber, userId, userIdSheet);
      if (selectionResult) {
        return selectionResult;
      }
      // ถ้าไม่ใช่การเลือก USER ID ให้ระบบอื่นจัดการต่อ
    }

    // ===== 5. จัดการ Postback =====
    if (requestObj.postback && requestObj.postback.data) {
      const postbackResult = handlePostbackEvents(token, requestObj.postback.data, userId, userIdSheet);
      if (postbackResult) {
        return postbackResult;
      }
      // ถ้าไม่ใช่ postback ของระบบ USER ID ให้ระบบอื่นจัดการต่อ
    }

    // ===== 6. ระบบบันทึกนัดหมาย =====
    if (typeof handleAppointmentBooking === 'function') {
      const appointmentResult = handleAppointmentBooking(userId, userMessage, token);
      if (appointmentResult) {
        console.log('📅 Appointment system handled the message');
        return appointmentResult;
      }
    }

    // ===== 7. ระบบปฏิทิน =====
    const calendarResult = handleCalendarSystem(requestObj, token, userMessage, userId);
    if (calendarResult) {
      console.log('📅 Calendar system handled the message');
      return calendarResult;
    }

    // ===== 8. รายงานรายเดือน =====
    if (userMessage.startsWith("สรุป")) {
      const requestedMonth = userMessage.replace("สรุป", "").trim();
      if (typeof MONTH_NAMES_TH !== 'undefined') {
        const monthIndex = MONTH_NAMES_TH.findIndex(month => month === requestedMonth);
        
        if (monthIndex !== -1) {
          const sheetName = "รายงาน" + requestedMonth;
          if (ss.getSheetByName(sheetName)) {
            const flexMessage = createMonthlyReport(sheetName);
            return replyFlexMessage(token, flexMessage);
          } else {
            return replyMessage(token, "ไม่พบรายงานเดือน" + requestedMonth);
          }
        }
      }
    }
    
    // Handle request for all monthly reports
    if (userMessage.toLowerCase() === "สรุป") {
      if (typeof sendAllMonthlyReports === 'function') {
        return sendAllMonthlyReports(token);
      }
    }

    // ===== 9. คำสั่งช่วยเหลือ =====
    if (userMessage && (userMessage.toLowerCase() === "รายชื่อผู้ใช้" || userMessage.toLowerCase() === "ผู้ใช้ทั้งหมด")) {
      return showAllUsers(token, userIdSheet);
    }
    
    // ===== 10. Keyword responses =====
    if (userMessage && keywordSheet && typeof getKeywordResponse === 'function') {
      const keywordResponse = getKeywordResponse(userMessage, keywordSheet);
      if (keywordResponse) {
        return replyMessage(token, keywordResponse);
      }
    }
    
    // If no specific action is matched, don't reply
    console.log('❓ No system handled the message, ignoring');
    return ContentService.createTextOutput(JSON.stringify({'status': 'ok'})).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('❌ Error in doPost:', error);
    return ContentService.createTextOutput(JSON.stringify({'status': 'error', 'message': error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ประมวลผลการค้นหา USER ID - ฟังก์ชันใหม่
 */
function processUserIdSearch(token, searchName, requesterId, userIdSheet) {
  try {
    console.log('🔍 Starting USER ID search for:', searchName, 'by:', requesterId);
    
    if (!searchName) {
      return replyMessage(token, "❌ กรุณาใส่ชื่อที่ต้องการค้นหา\nตัวอย่าง: ID กมลภพ");
    }

    const lastRow = userIdSheet.getLastRow();
    console.log('📊 Total rows in sheet:', lastRow);
    
    if (lastRow < 2) {
      return replyMessage(token, "❌ ไม่พบข้อมูลผู้ใช้ในระบบ");
    }

    // ดึงข้อมูลทั้งหมดจากชีต (คอลัมน์ A = USER ID, คอลัมน์ B = ชื่อ)
    const data = userIdSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    console.log('📋 Data retrieved:', data.length, 'rows');
    
    const matches = findMatchingUsers(data, searchName);
    console.log('📊 Total matches found:', matches.length);
    
    if (matches.length === 0) {
      return showNoUserFoundMessage(token, searchName, data);
    }

    if (matches.length === 1) {
      // พบผู้ใช้เพียงคนเดียว - บันทึกทันที
      return processSingleUserFound(token, matches[0], requesterId, userIdSheet);
    }

    // พบผู้ใช้หลายคน - แสดงรายการให้เลือก
    return showUserSelectionList(token, matches, searchName, requesterId);
    
  } catch (error) {
    console.error('❌ Error in processUserIdSearch:', error);
    return replyMessage(token, "⚠️ เกิดข้อผิดพลาดในการค้นหา: " + error.toString());
  }
}

/**
 * ค้นหาผู้ใช้ที่ตรงกัน - ฟังก์ชันใหม่
 */
function findMatchingUsers(data, searchName) {
  const exactMatches = [];
  const partialMatches = [];
  
  for (let i = 0; i < data.length; i++) {
    const userId = data[i][0];
    const fullName = data[i][1];
    
    // ตรวจสอบว่ามี USER ID และชื่อ และไม่ใช่ Group
    if (userId && fullName && fullName.toString().trim() !== "" && fullName.toString() !== "Group") {
      const nameStr = fullName.toString().trim();
      const searchStr = searchName.trim();
      
      console.log(`📝 Checking: "${nameStr}" against "${searchStr}"`);
      
      // ตรวจสอบการตรงกันแบบต่างๆ
      if (nameStr === searchStr || nameStr.toLowerCase() === searchStr.toLowerCase()) {
        // ตรงกัน 100%
        exactMatches.push({
          userId: userId.toString(),
          fullName: nameStr,
          matchType: 'exact',
          rowIndex: i + 2
        });
        console.log('✅ Exact match found:', nameStr);
      }
      else if (nameStr.toLowerCase().includes(searchStr.toLowerCase()) || 
               searchStr.toLowerCase().includes(nameStr.toLowerCase())) {
        // มีชื่อที่ค้นหาอยู่ในชื่อเต็ม
        partialMatches.push({
          userId: userId.toString(),
          fullName: nameStr,
          matchType: 'contains',
          rowIndex: i + 2
        });
        console.log('🔍 Partial match found:', nameStr);
      }
      else {
        // ตรวจสอบคำแรกของชื่อ (เช่น "กมลภพ" จะเจอ "กมลภพ จำปา", "กมลภพ จำปี")
        const nameWords = nameStr.toLowerCase().split(/\s+/);
        const searchWords = searchStr.toLowerCase().split(/\s+/);
        
        let hasCommonWord = false;
        for (const searchWord of searchWords) {
          for (const nameWord of nameWords) {
            if (searchWord && nameWord && 
                (nameWord.includes(searchWord) || searchWord.includes(nameWord) || 
                 nameWord.startsWith(searchWord) || searchWord.startsWith(nameWord))) {
              hasCommonWord = true;
              break;
            }
          }
          if (hasCommonWord) break;
        }
        
        if (hasCommonWord) {
          partialMatches.push({
            userId: userId.toString(),
            fullName: nameStr,
            matchType: 'word_match',
            rowIndex: i + 2
          });
          console.log('📝 Word match found:', nameStr);
        }
      }
    }
  }

  // รวมผลลัพธ์ โดยให้ exact matches มาก่อน
  const allMatches = [...exactMatches, ...partialMatches];
  
  // เรียงลำดับ: exact matches ก่อน
  allMatches.sort((a, b) => {
    if (a.matchType === 'exact' && b.matchType !== 'exact') return -1;
    if (a.matchType !== 'exact' && b.matchType === 'exact') return 1;
    return 0;
  });

  return allMatches;
}

/**
 * แสดงข้อความเมื่อไม่พบผู้ใช้ - ฟังก์ชันใหม่
 */
function showNoUserFoundMessage(token, searchName, data) {
  let message = `❌ ไม่พบผู้ใช้ที่มีชื่อตรงกับ "${searchName}"\n\n`;
  
  // แสดงชื่อผู้ใช้ในระบบ (5 คนแรก)
  const allNames = data.filter(row => row[1] && row[1].toString().trim() !== "" && row[1].toString() !== "Group")
                       .map(row => row[1].toString().trim())
                       .slice(0, 5);
  
  if (allNames.length > 0) {
    message += `💡 ชื่อผู้ใช้ในระบบ (${allNames.length > 5 ? '5 คนแรก' : allNames.length + ' คน'}):\n`;
    allNames.forEach((name, index) => {
      message += `${index + 1}. ${name}\n`;
    });
    message += `\n📝 ลองใช้คำว่า: ID ${allNames[0]}`;
  } else {
    message += `💡 ไม่พบข้อมูลผู้ใช้ในระบบ`;
  }
  
  return replyMessage(token, message);
}

/**
 * ประมวลผลเมื่อพบผู้ใช้เพียงคนเดียว - ฟังก์ชันใหม่
 */
function processSingleUserFound(token, match, requesterId, userIdSheet) {
  console.log('✅ Single user found:', match.fullName, match.userId);
  
  // บันทึก USER ID ลงคอลัมน์ G
  recordUserIdInColumnG(userIdSheet, requesterId, match.userId);
  
  // เรียกใช้การบันทึกอัตโนมัติ (สำหรับการนัดหมาย)
  if (typeof handleFoundUserId === 'function') {
    const autoSaveResult = handleFoundUserId(requesterId, match.userId, token);
    if (autoSaveResult) {
      return autoSaveResult;
    }
  }
  
  return replyMessage(token, 
    `✅ พบผู้ใช้: ${match.fullName}\n` +
    `🆔 USER ID: ${match.userId}\n\n` +
    `✅ บันทึก USER ID เรียบร้อยแล้ว`
  );
}

/**
 * แสดงรายการให้เลือกเมื่อพบผู้ใช้หลายคน - ฟังก์ชันใหม่
 */
function showUserSelectionList(token, matches, searchName, requesterId) {
  try {
    console.log('📋 Creating selection list for', matches.length, 'users');
    
    // จำกัดให้แสดงแค่ 9 คนแรก
    const displayMatches = matches.slice(0, 9);
    
    // สร้างข้อความแสดงรายการ
    let messageText = `🔍 พบผู้ใช้ ${matches.length} คนที่มีชื่อคล้ายกับ "${searchName}"\n\n`;
    messageText += `📋 รายชื่อที่พบ:\n`;
    
    displayMatches.forEach((match, index) => {
      const icon = match.matchType === 'exact' ? '🎯' : '🔍';
      messageText += `${index + 1}. ${icon} ${match.fullName}\n`;
    });
    
    if (matches.length > 9) {
      messageText += `\n... และอีก ${matches.length - 9} คน`;
    }
    
    messageText += `\n\n👆 กรุณาเลือกด้วยตัวเลือกด้านล่าง`;

    // บันทึกรายการสำหรับการเลือก
    const options = displayMatches.map((match, index) => ({
      userId: match.userId,
      fullName: match.fullName,
      number: index + 1
    }));
    
    storeUserSelectionOptions(requesterId, options);

    // สร้าง Quick Reply
    const quickReplyItems = displayMatches.map((match, index) => {
      let displayName = match.fullName;
      if (displayName.length > 15) {
        displayName = displayName.substring(0, 12) + "...";
      }
      
      return {
        type: "action",
        action: {
          type: "message",
          label: `${index + 1}. ${displayName}`,
          text: `${index + 1}`
        }
      };
    });

    // ส่งข้อความพร้อม Quick Reply
    const replyData = {
      replyToken: token,
      messages: [{
        type: "text",
        text: messageText,
        quickReply: {
          items: quickReplyItems
        }
      }]
    };

    console.log('📤 Sending Quick Reply with', quickReplyItems.length, 'options');

    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
      },
      method: 'POST',
      payload: JSON.stringify(replyData),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    console.log('📨 Quick Reply response code:', responseCode);

    if (responseCode >= 200 && responseCode < 300) {
      console.log('✅ Quick Reply sent successfully');
      return ContentService.createTextOutput(JSON.stringify({'status': 'ok'})).setMimeType(ContentService.MimeType.JSON);
    } else {
      console.log('❌ Quick Reply failed, response:', response.getContentText());
      
      // Fallback: ส่งข้อความธรรมดา
      const fallbackMessage = messageText + `\n\n📝 พิมพ์หมายเลข 1-${displayMatches.length} เพื่อเลือก`;
      return replyMessage(token, fallbackMessage);
    }
    
  } catch (error) {
    console.error('❌ Error in showUserSelectionList:', error);
    
    // Ultimate fallback
    const simpleMessage = `🔍 พบผู้ใช้ ${matches.length} คน:\n\n` +
                         matches.slice(0, 5).map((match, i) => `${i + 1}. ${match.fullName}`).join('\n') +
                         `\n\n📝 พิมพ์หมายเลข 1-${Math.min(matches.length, 5)} เพื่อเลือก`;
    
    return replyMessage(token, simpleMessage);
  }
}

/**
 * ประมวลผลการเลือกด้วยหมายเลข - ฟังก์ชันใหม่
 */
function processNumberSelection(token, selectedNumber, requesterId, userIdSheet) {
  try {
    console.log('🔢 Processing number selection:', selectedNumber, 'by:', requesterId);
    
    // ดึงรายการที่เก็บไว้
    const storedOptions = getUserSelectionOptions(requesterId);
    
    if (!storedOptions || storedOptions.length === 0) {
      console.log('❌ No stored selection options found');
      return null; // ให้ระบบอื่นจัดการต่อ (เช่น ระบบปฏิทิน)
    }
    
    console.log('📋 Found', storedOptions.length, 'stored options');
    
    if (selectedNumber < 1 || selectedNumber > storedOptions.length) {
      return replyMessage(token, `❌ กรุณาเลือกหมายเลข 1-${storedOptions.length}`);
    }
    
    const selectedOption = storedOptions[selectedNumber - 1];
    const selectedUserId = selectedOption.userId;
    const selectedName = selectedOption.fullName;
    
    console.log('👤 User selected:', selectedName, 'ID:', selectedUserId);
    
    // บันทึก USER ID ลงคอลัมน์ G
    recordUserIdInColumnG(userIdSheet, requesterId, selectedUserId);
    
    // ลบรายการที่เก็บไว้
    clearUserSelectionOptions(requesterId);
    
    // เรียกใช้การบันทึกอัตโนมัติ (สำหรับการนัดหมาย)
    if (typeof handleFoundUserId === 'function') {
      const autoSaveResult = handleFoundUserId(requesterId, selectedUserId, token);
      if (autoSaveResult) {
        return autoSaveResult;
      }
    }
    
    return replyMessage(token, 
      `✅ เลือกผู้ใช้: ${selectedName}\n` +
      `🆔 USER ID: ${selectedUserId}\n\n` +
      `✅ บันทึก USER ID เรียบร้อยแล้ว`
    );
    
  } catch (error) {
    console.error('❌ Error in processNumberSelection:', error);
    return replyMessage(token, "⚠️ เกิดข้อผิดพลาดในการเลือกผู้ใช้");
  }
}

/**
 * จัดการ Postback Events - ฟังก์ชันใหม่
 */
function handlePostbackEvents(token, postbackData, requesterId, userIdSheet) {
  try {
    console.log('📲 Handling postback:', postbackData);
    
    // จัดการ postback สำหรับการเลือกผู้ใช้
    if (postbackData.startsWith("select_user_")) {
      const selectedUserId = postbackData.replace("select_user_", "");
      
      // หาข้อมูลผู้ใช้ที่เลือก
      const lastRow = userIdSheet.getLastRow();
      const data = userIdSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString() === selectedUserId) {
          const fullName = data[i][1] ? data[i][1].toString() : "ไม่ระบุชื่อ";
          
          // บันทึก USER ID
          recordUserIdInColumnG(userIdSheet, requesterId, selectedUserId);
          
          // ลบรายการที่เก็บไว้
          clearUserSelectionOptions(requesterId);
          
          // เรียกใช้การบันทึกอัตโนมัติ
          if (typeof handleFoundUserId === 'function') {
            const autoSaveResult = handleFoundUserId(requesterId, selectedUserId, token);
            if (autoSaveResult) {
              return autoSaveResult;
            }
          }
          
          return replyMessage(token, 
            `✅ เลือกผู้ใช้: ${fullName}\n` +
            `🆔 USER ID: ${selectedUserId}\n\n` +
            `✅ บันทึก USER ID เรียบร้อยแล้ว`
          );
        }
      }
      
      return replyMessage(token, "❌ ไม่พบข้อมูลผู้ใช้ที่เลือก");
    }
    
    // จัดการ postback อื่นๆ (เช่น ปฏิทิน)
    return handleCalendarPostback(postbackData, token, requesterId);
    
  } catch (error) {
    console.error('❌ Error in handlePostbackEvents:', error);
    return null;
  }
}

/**
 * จัดการ postback ของปฏิทิน - ฟังก์ชันใหม่
 */
function handleCalendarPostback(postbackData, token, userId) {
  try {
    if (postbackData.indexOf('prev_month_') === 0) {
      const parts = postbackData.split('_');
      const year = parseInt(parts[2]);
      const month = parseInt(parts[3]);
      const prevDate = new Date(year, month - 2, 1);
      const message = createCalendarFlexMessage(prevDate.getFullYear(), prevDate.getMonth() + 1);
      return replyFlexMessage(token, message);
    }
    
    if (postbackData.indexOf('next_month_') === 0) {
      const parts = postbackData.split('_');
      const year = parseInt(parts[2]);
      const month = parseInt(parts[3]);
      const nextDate = new Date(year, month, 1);
      const message = createCalendarFlexMessage(nextDate.getFullYear(), nextDate.getMonth() + 1);
      return replyFlexMessage(token, message);
    }
    
    if (postbackData.indexOf('select_date_') === 0) {
      const parts = postbackData.split('_');
      const year = parts[2];
      const month = parts[3];
      const day = parts[4];
      
      const selectedDate = year + '-' + padZero(month) + '-' + padZero(day);
      const displayDate = day + '/' + month + '/' + year;
      
      const calendarEvents = getDetailedCalendarEventsForDate(selectedDate);
      
      if (calendarEvents.length > 0) {
        const message = createEventDetailsFlexMessage(displayDate, calendarEvents);
        return replyFlexMessage(token, message);
      } else {
        const message = {
          type: 'text',
          text: '📅 วันที่: ' + displayDate + '\n\n❌ ไม่มีกิจกรรมในวันนี้'
        };
        return replyMessage(token, message);
      }
    }

    return null;
    
  } catch (error) {
    console.error('❌ Error in handleCalendarPostback:', error);
    return null;
  }
}

/**
 * ระบบปฏิทิน - ฟังก์ชันใหม่
 */
function handleCalendarSystem(requestObj, token, userMessage, userId) {
  try {
    // ตรวจสอบคำสั่งปฏิทินพื้นฐาน
    if (userMessage === 'ปฏิทิน' || userMessage === 'calendar') {
      const today = new Date();
      const message = createCalendarFlexMessage(today.getFullYear(), today.getMonth() + 1);
      return replyFlexMessage(token, message);
    }
    
    if (userMessage === 'เมนู' || userMessage === 'menu') {
      const message = {
        type: 'text',
        text: '📋 เมนูคำสั่ง:\n\n' +
              '📅 ปฏิทิน - แสดงปฏิทินพร้อมข้อมูลจาก Google Calendar\n' +
              '❓ help - วิธีใช้งาน\n' +
              '🔢 1-31 - เลือกวันที่ในเดือนปัจจุบัน\n' +
              '📝 ลงเวลานัดหมาย - เริ่มบันทึกข้อมูลนัดหมาย\n\n' +
              'พิมพ์คำสั่งที่ต้องการได้เลยครับ!'
      };
      return replyMessage(token, message);
    }
    
    if (userMessage.toLowerCase() === 'help' || userMessage === 'ช่วยเหลือ') {
      const message = {
        type: 'text',
        text: '📅 วิธีใช้งานปฏิทิน:\n\n' +
              '🔹 พิมพ์ "ปฏิทิน" - แสดงปฏิทินพร้อมข้อมูลจาก Google Calendar\n' +
              '🔹 พิมพ์ "เมนู" - แสดงรายการคำสั่ง\n' +
              '🔹 พิมพ์ตัวเลข 1-31 - เลือกวันที่ในเดือนปัจจุบัน\n' +
              '🔹 พิมพ์ "ลงเวลานัดหมาย" - เริ่มบันทึกข้อมูลนัดหมาย\n\n' +
              '📋 ดูกิจกรรมตามวันที่:\n' +
              '• 12/7/2568 - ดูกิจกรรมวันที่ 12 กรกฎาคม 2568\n' +
              '• 12 กรกฎาคม 2568 - ดูกิจกรรมวันที่เฉพาะ\n\n' +
              '📋 ดูกิจกรรมทั้งเดือน:\n' +
              '• 7/2568 - ดูกิจกรรมทั้งเดือนกรกฎาคม 2568\n' +
              '• กรกฎาคม 2568 - ดูกิจกรรมทั้งเดือน\n\n' +
              '✨ วันที่มีกิจกรรมจะแสดงสีพิเศษในปฏิทิน\n' +
              'ข้อมูลจะบันทึกลง Google Sheets อัตโนมัติ'
      };
      return replyMessage(token, message);
    }
    
    if (userMessage === 'test' || userMessage === 'ทดสอบ') {
      const message = {
        type: 'text',
        text: '✅ บอททำงานปกติ!\nเวลา: ' + new Date().toLocaleString('th-TH') + '\nUser ID: ' + userId
      };
      return replyMessage(token, message);
    }

    // จัดการวันที่แบบลอยๆ สำหรับปฏิทิน
    if (userMessage && isDateFormatForCalendar(userMessage)) {
      return handleCalendarDateQuery(userMessage, token, userId);
    }

    // เช็คตัวเลข (สำหรับเลือกวันที่ในปฏิทิน) - แต่ต้องแน่ใจว่าไม่ใช่การเลือก USER ID
    if (/^\d{1,2}$/.test(userMessage)) {
      // ตรวจสอบว่าผู้ใช้มีรายการเลือก USER ID หรือไม่
      const storedOptions = getUserSelectionOptions(userId);
      if (storedOptions && storedOptions.length > 0) {
        // ถ้ามีรายการเลือก USER ID ให้ระบบนั้นจัดการ (return null)
        return null;
      }
      
      // ถ้าไม่มีรายการเลือก USER ID ให้ระบบปฏิทินจัดการ
      return handleCalendarDaySelection(userMessage, token, userId);
    }

    return null; // ไม่ใช่คำสั่งปฏิทิน
    
  } catch (error) {
    console.error('❌ Error in handleCalendarSystem:', error);
    return null;
  }
}

/**
 * จัดการการเลือกวันที่ในปฏิทิน - ฟังก์ชันใหม่
 */
function handleCalendarDaySelection(userMessage, token, userId) {
  try {
    const day = parseInt(userMessage);
    const today = new Date();
    
    if (day >= 1 && day <= 31) {
      const year = today.getFullYear();
      const month = today.getMonth() + 1;
      const daysInCurrentMonth = new Date(year, month, 0).getDate();
      
      if (day <= daysInCurrentMonth) {
        const selectedDate = year + '-' + padZero(month) + '-' + padZero(day);
        const displayDate = day + '/' + month + '/' + year;
        
        const calendarEvents = getDetailedCalendarEventsForDate(selectedDate);
        
        if (calendarEvents.length > 0) {
          const message = createEventDetailsFlexMessage(displayDate, calendarEvents);
          return replyFlexMessage(token, message);
        } else {
          const message = {
            type: 'text',
            text: '📅 วันที่: ' + displayDate + '\n\n❌ ไม่มีกิจกรรมในวันนี้'
          };
          return replyMessage(token, message);
        }
      } else {
        const message = {
          type: 'text',
          text: 'วันที่ ' + day + ' ไม่มีในเดือนนี้ กรุณาเลือกใหม่'
        };
        return replyMessage(token, message);
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ Error in handleCalendarDaySelection:', error);
    return null;
  }
}

/**
 * จัดการคำถามเกี่ยวกับวันที่แบบลอยๆ สำหรับปฏิทิน - ฟังก์ชันใหม่
 */
function handleCalendarDateQuery(userMessage, token, userId) {
  try {
    const parsedDate = parseThaiDateForCalendar(userMessage);
    
    if (parsedDate) {
      const year = parsedDate.year;
      const month = parsedDate.month;
      const day = parsedDate.day;
      
      if (day === 0) {
        // แสดงกิจกรรมทั้งเดือน
        const message = createMonthlyEventsFlexMessage(month, year);
        return replyFlexMessage(token, message);
      } else {
        // แสดงกิจกรรมวันเฉพาะ
        const selectedDate = year + '-' + padZero(month) + '-' + padZero(day);
        const displayDate = day + '/' + month + '/' + year;
        
        const calendarEvents = getDetailedCalendarEventsForDate(selectedDate);
        
        if (calendarEvents.length > 0) {
          const message = createEventDetailsFlexMessage(displayDate, calendarEvents);
          return replyFlexMessage(token, message);
        } else {
          const message = {
            type: 'text',
            text: '📅 วันที่: ' + displayDate + '\n\n❌ ไม่มีกิจกรรมในวันนี้'
          };
          return replyMessage(token, message);
        }
      }
    } else {
      const message = {
        type: 'text',
        text: '❌ รูปแบบวันที่ไม่ถูกต้อง\n\nรูปแบบที่รองรับ:\n' +
              '• 12/7/2568\n' +
              '• 12 กรกฎาคม 2568\n' +
              '• 7/2568 (ดูกิจกรรมทั้งเดือน)\n' +
              '• กรกฎาคม 2568 (ดูกิจกรรมทั้งเดือน)'
      };
      return replyMessage(token, message);
    }
    
  } catch (error) {
    console.error('❌ Error in handleCalendarDateQuery:', error);
    return null;
  }
}

/**
 * ตรวจสอบรูปแบบวันที่สำหรับปฏิทิน - ฟังก์ชันใหม่
 */
function isDateFormatForCalendar(text) {
  const datePattern1 = /^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}$/;
  const datePattern2 = /^\d{1,2}[\/\-\.]\d{4}$/;
  const thaiDatePattern = /^\d{1,2}\s+[ก-๙]+\s+\d{4}$/;
  const thaiMonthPattern = /^[ก-๙]+\s+\d{4}$/;
  
  return datePattern1.test(text) || datePattern2.test(text) || 
         thaiDatePattern.test(text) || thaiMonthPattern.test(text);
}

/**
 * แปลงวันที่ไทยสำหรับปฏิทิน - ฟังก์ชันใหม่
 */
function parseThaiDateForCalendar(text) {
  const thaiMonths = {
    'มกราคม': 1, 'มค': 1, 'ม.ค.': 1,
    'กุมภาพันธ์': 2, 'กพ': 2, 'ก.พ.': 2,
    'มีนาคม': 3, 'มีค': 3, 'มี.ค.': 3,
    'เมษายน': 4, 'เมย': 4, 'เม.ย.': 4,
    'พฤษภาคม': 5, 'พค': 5, 'พ.ค.': 5,
    'มิถุนายน': 6, 'มิย': 6, 'มิ.ย.': 6,
    'กรกฎาคม': 7, 'กค': 7, 'ก.ค.': 7,
    'สิงหาคม': 8, 'สค': 8, 'ส.ค.': 8,
    'กันยายน': 9, 'กย': 9, 'ก.ย.': 9,
    'ตุลาคม': 10, 'ตค': 10, 'ต.ค.': 10,
    'พฤศจิกายน': 11, 'พย': 11, 'พ.ย.': 11,
    'ธันวาคม': 12, 'ธค': 12, 'ธ.ค.': 12
  };
  
  try {
    // รูปแบบ: 12/7/2568
    const match1 = text.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
    if (match1) {
      const day = parseInt(match1[1]);
      const month = parseInt(match1[2]);
      const year = parseInt(match1[3]);
      const gregorianYear = year > 2500 ? year - 543 : year;
      
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return { day: day, month: month, year: gregorianYear };
      }
    }
    
    // รูปแบบ: 7/2568 (เดือน/ปี)
    const match2 = text.match(/^(\d{1,2})[\/\-\.](\d{4})$/);
    if (match2) {
      const month = parseInt(match2[1]);
      const year = parseInt(match2[2]);
      const gregorianYear = year > 2500 ? year - 543 : year;
      
      if (month >= 1 && month <= 12) {
        return { day: 0, month: month, year: gregorianYear };
      }
    }
    
    // รูปแบบ: 12 กรกฎาคม 2568
    const match3 = text.match(/^(\d{1,2})\s+([ก-๙]+)\s+(\d{4})$/);
    if (match3) {
      const day = parseInt(match3[1]);
      const monthName = match3[2];
      const year = parseInt(match3[3]);
      const gregorianYear = year > 2500 ? year - 543 : year;
      
      const month = thaiMonths[monthName];
      if (month && day >= 1 && day <= 31) {
        return { day: day, month: month, year: gregorianYear };
      }
    }
    
    // รูปแบบ: กรกฎาคม 2568 (เดือน ปี)
    const match4 = text.match(/^([ก-๙]+)\s+(\d{4})$/);
    if (match4) {
      const monthName = match4[1];
      const year = parseInt(match4[2]);
      const gregorianYear = year > 2500 ? year - 543 : year;
      
      const month = thaiMonths[monthName];
      if (month) {
        return { day: 0, month: month, year: gregorianYear };
      }
    }
    
    return null;
  } catch (error) {
    console.error('Error parsing Thai date for calendar:', error);
    return null;
  }
}

// ===== ฟังก์ชันจัดการข้อมูล USER ID =====

/**
 * เก็บรายการเลือกผู้ใช้ - ฟังก์ชันใหม่
 */
function storeUserSelectionOptions(requesterId, options) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const key = `user_selection_${requesterId}`;
    
    const data = {
      options: options,
      timestamp: new Date().getTime()
    };
    
    properties.setProperty(key, JSON.stringify(data));
    console.log('💾 Stored user selection options for:', requesterId, 'count:', options.length);
    
  } catch (error) {
    console.error('❌ Error storing user selection options:', error);
  }
}

/**
 * ดึงรายการเลือกผู้ใช้ - ฟังก์ชันใหม่
 */
function getUserSelectionOptions(requesterId) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const key = `user_selection_${requesterId}`;
    const storedData = properties.getProperty(key);
    
    if (!storedData) {
      return null;
    }
    
    const data = JSON.parse(storedData);
    
    // ตรวจสอบเวลา (หมดอายุใน 10 นาที)
    const now = new Date().getTime();
    const timeDiff = now - data.timestamp;
    
    if (timeDiff > 600000) { // 10 minutes
      properties.deleteProperty(key);
      console.log('⌛ User selection options expired for:', requesterId);
      return null;
    }
    
    console.log('📋 Retrieved user selection options for:', requesterId, 'count:', data.options.length);
    return data.options;
    
  } catch (error) {
    console.error('❌ Error getting user selection options:', error);
    return null;
  }
}

/**
 * ลบรายการเลือกผู้ใช้ - ฟังก์ชันใหม่
 */
function clearUserSelectionOptions(requesterId) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const key = `user_selection_${requesterId}`;
    properties.deleteProperty(key);
    console.log('🗑️ Cleared user selection options for:', requesterId);
  } catch (error) {
    console.error('❌ Error clearing user selection options:', error);
  }
}

/**
 * บันทึก USER ID ลงคอลัมน์ G - ฟังก์ชันใหม่
 */
function recordUserIdInColumnG(userIdSheet, requesterId, targetUserId) {
  try {
    console.log('💾 Recording USER ID - Requester:', requesterId, 'Target:', targetUserId);
    
    const lastRow = userIdSheet.getLastRow();
    
    // หาแถวของผู้ที่ขอข้อมูล
    for (let i = 2; i <= lastRow; i++) {
      const userId = userIdSheet.getRange(i, 1).getValue();
      if (userId && userId.toString() === requesterId.toString()) {
        // บันทึก USER ID ที่ค้นหาได้ลงคอลัมน์ G
        userIdSheet.getRange(i, 7).setValue(targetUserId);
        
        // บันทึกเวลาที่ค้นหาลงคอลัมน์ H
        const currentTime = new Date();
        userIdSheet.getRange(i, 8).setValue(currentTime);
        
        console.log('✅ Recorded USER ID in row', i, 'column G:', targetUserId);
        break;
      }
    }
    
  } catch (error) {
    console.error('❌ Error in recordUserIdInColumnG:', error);
  }
}

// ===== ฟังก์ชันสนับสนุน =====

/**
 * ดึงข้อมูลโปรไฟล์ของผู้ใช้จาก LINE API - ฟังก์ชันใหม่
 */
function getUserProfile(userId) {
  const url = `https://api.line.me/v2/bot/profile/${userId}`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${CHANNEL_ACCESS_TOKEN}`
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const profile = JSON.parse(response.getContentText());
      console.log('✅ ดึงข้อมูลโปรไฟล์สำเร็จ:', profile.displayName);
      return profile;
    } else {
      console.log('❌ Error fetching profile:', response.getContentText());
      return null;
    }
  } catch (error) {
    console.log('❌ Fetch Error:', error);
    return null;
  }
}

/**
 * แสดงรายชื่อผู้ใช้ทั้งหมดในระบบ - ฟังก์ชันใหม่
 */
function showAllUsers(token, userIdSheet) {
  try {
    const lastRow = userIdSheet.getLastRow();
    
    if (lastRow < 2) {
      return replyMessage(token, "❌ ไม่พบข้อมูลผู้ใช้ในระบบ");
    }
    
    // ดึงข้อมูลชื่อผู้ใช้ทั้งหมด
    const data = userIdSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const validUsers = data.filter(row => 
      row[1] && 
      row[1].toString().trim() !== "" && 
      row[1].toString() !== "Group"
    );
    
    if (validUsers.length === 0) {
      return replyMessage(token, "❌ ไม่พบข้อมูลผู้ใช้ที่ถูกต้องในระบบ");
    }
    
    // จำกัดให้แสดงแค่ 20 คนแรก
    const displayUsers = validUsers.slice(0, 20);
    
    let message = `👥 รายชื่อผู้ใช้ในระบบ (${validUsers.length} คน):\n\n`;
    
    displayUsers.forEach((row, index) => {
      const name = row[1].toString().trim();
      message += `${index + 1}. ${name}\n`;
    });
    
    if (validUsers.length > 20) {
      message += `\n... และอีก ${validUsers.length - 20} คน`;
    }
    
    message += `\n\n💡 ใช้คำสั่ง: ID ชื่อ\nตัวอย่าง: ID ${displayUsers[0][1]}`;
    
    return replyMessage(token, message);
    
  } catch (error) {
    console.error('❌ Error in showAllUsers:', error);
    return replyMessage(token, "⚠️ เกิดข้อผิดพลาดในการดึงรายชื่อผู้ใช้: " + error.toString());
  }
}

/**
 * ฟังก์ชันสำหรับทดสอบ GET request
 */
function doGet(e) {
  return ContentService.createTextOutput('LINE Bot is running! Time: ' + new Date().toLocaleString('th-TH')).setMimeType(ContentService.MimeType.TEXT);
}

// ===== ฟังก์ชันระบบปฏิทิน =====

/**
 * สร้าง Calendar Flex Message พร้อมข้อมูลจาก Google Calendar
 */
function createCalendarFlexMessage(year, month) {
  var monthNames = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  
  var firstDay = new Date(year, month - 1, 1).getDay();
  var daysInMonth = new Date(year, month, 0).getDate();
  
  // ดึงข้อมูลกิจกรรมจาก Google Calendar สำหรับเดือนนี้
  var monthlyEvents = getCalendarEventsForMonth(year, month);
  
  // Header ของปฏิทิน
  var header = {
    type: "box",
    layout: "horizontal",
    contents: [
      {
        type: "text",
        text: "◀",
        action: {
          type: "postback",
          data: "prev_month_" + year + "_" + month
        },
        flex: 1,
        align: "center",
        color: "#0066CC",
        size: "lg",
        weight: "bold"
      },
      {
        type: "text",
        text: monthNames[month - 1] + " " + year,
        weight: "bold",
        size: "lg",
        align: "center",
        flex: 3
      },
      {
        type: "text",
        text: "▶",
        action: {
          type: "postback",
          data: "next_month_" + year + "_" + month
        },
        flex: 1,
        align: "center",
        color: "#0066CC",
        size: "lg",
        weight: "bold"
      }
    ],
    paddingBottom: "md"
  };
  
  // วันในสัปดาห์
  var dayHeaders = {
    type: "box",
    layout: "horizontal",
    contents: [
      createDayHeader('อา'),
      createDayHeader('จ'),
      createDayHeader('อ'),
      createDayHeader('พ'),
      createDayHeader('พฤ'),
      createDayHeader('ศ'),
      createDayHeader('ส')
    ],
    paddingBottom: "sm"
  };
  
  // สร้างตารางวันที่
  var weeks = [];
  var currentWeek = [];
  
  // เพิ่มช่องว่างสำหรับวันแรก
  for (var i = 0; i < firstDay; i++) {
    currentWeek.push(createEmptyCell());
  }
  
  // เพิ่มวันที่
  for (var day = 1; day <= daysInMonth; day++) {
    var hasEvent = monthlyEvents.hasOwnProperty(day);
    currentWeek.push(createDayCell(day, year, month, hasEvent));
    
    // ถ้าครบ 7 วัน (หนึ่งสัปดาห์)
    if (currentWeek.length === 7) {
      weeks.push({
        type: "box",
        layout: "horizontal",
        contents: currentWeek.slice(),
        paddingTop: "xs",
        paddingBottom: "xs"
      });
      currentWeek = [];
    }
  }
  
  // เพิ่มช่องว่างท้ายสัปดาห์สุดท้าย
  while (currentWeek.length < 7 && currentWeek.length > 0) {
    currentWeek.push(createEmptyCell());
  }
  
  if (currentWeek.length > 0) {
    weeks.push({
      type: "box",
      layout: "horizontal",
      contents: currentWeek.slice(),
      paddingTop: "xs",
      paddingBottom: "xs"
    });
  }
  
  return {
    type: "flex",
    altText: "เลือกวันที่",
    contents: {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        contents: [header, dayHeaders].concat(weeks),
        spacing: "none",
        paddingAll: "lg"
      }
    }
  };
}

/**
 * สร้าง header วัน
 */
function createDayHeader(day) {
  return {
    type: "text",
    text: day,
    size: "sm",
    align: "center",
    weight: "bold",
    flex: 1,
    color: "#999999"
  };
}

/**
 * สร้างช่องวันที่ที่กดได้ (แสดงตัวเลขชัดเจน)
 */
function createDayCell(day, year, month, hasEvent) {
  var textColor = "#0066CC";
  var backgroundColor = "#FFFFFF";
  
  // ถ้ามีกิจกรรมในวันนี้ ให้เปลี่ยนสี
  if (hasEvent) {
    textColor = "#FFFFFF";
    backgroundColor = "#FF6B6B"; // สีแดงอ่อนสำหรับวันที่มีกิจกรรม
  }
  
  // เช็คว่าเป็นวันปัจจุบันหรือไม่
  var today = new Date();
  var isToday = (today.getFullYear() === year && 
                 today.getMonth() + 1 === month && 
                 today.getDate() === day);
  
  if (isToday && !hasEvent) {
    backgroundColor = "#E3F2FD"; // สีฟ้าอ่อนสำหรับวันปัจจุบัน
  }
  
  return {
    type: "text",
    text: day.toString(),
    size: "md",
    align: "center",
    flex: 1,
    action: {
      type: "postback",
      data: "select_date_" + year + "_" + month + "_" + day
    },
    color: textColor,
    weight: isToday ? "bold" : "regular",
    backgroundColor: backgroundColor
  };
}

/**
 * สร้างช่องว่าง
 */
function createEmptyCell() {
  return {
    type: "text",
    text: "　", // ใช้ full-width space
    size: "md",
    align: "center",
    flex: 1,
    color: "#FFFFFF"
  };
}

/**
 * ดึงข้อมูลกิจกรรมจาก Google Calendar สำหรับเดือนหนึ่ง
 */
function getCalendarEventsForMonth(year, month) {
  try {
    var calendar = CalendarApp.getCalendarById(APPOINTMENT_CALENDAR_ID);
    var startDate = new Date(year, month - 1, 1);
    var endDate = new Date(year, month, 0, 23, 59, 59);
    
    var events = calendar.getEvents(startDate, endDate);
    var eventsMap = {};
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var eventDate = event.getStartTime();
      var day = eventDate.getDate();
      
      if (!eventsMap[day]) {
        eventsMap[day] = [];
      }
      eventsMap[day].push(event.getTitle());
    }
    
    return eventsMap;
  } catch (error) {
    console.error('Error getting calendar events for month:', error);
    return {};
  }
}

/**
 * ดึงข้อมูลกิจกรรมจาก Google Calendar สำหรับวันที่เฉพาะ (รายละเอียดครบถ้วน)
 */
function getDetailedCalendarEventsForDate(dateString) {
  try {
    var calendar = CalendarApp.getCalendarById(APPOINTMENT_CALENDAR_ID);
    var date = new Date(dateString);
    var startDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    var endDate = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
    
    var events = calendar.getEvents(startDate, endDate);
    var eventDetails = [];
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var eventDetail = {
        title: event.getTitle(),
        description: event.getDescription() || '',
        location: event.getLocation() || '',
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        isAllDay: event.isAllDayEvent(),
        guests: event.getGuestList().map(function(guest) { return guest.getEmail(); }),
        creator: event.getCreators()[0] || '',
        color: event.getColor() || CalendarApp.EventColor.BLUE
      };
      
      eventDetails.push(eventDetail);
    }
    
    return eventDetails;
  } catch (error) {
    console.error('Error getting detailed calendar events for date:', error);
    return [];
  }
}

/**
 * สร้าง Flex Message แสดงรายละเอียดกิจกรรม
 */
function createEventDetailsFlexMessage(displayDate, events) {
  var bubbles = [];
  
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    
    // กำหนดสีหัวข้อตามประเภทกิจกรรม
    var headerColor = getEventHeaderColor(event.title);
    
    // สร้างข้อความเวลา
    var timeText = "";
    if (event.isAllDay) {
      timeText = "ตลอดวัน";
    } else {
      timeText = Utilities.formatDate(event.startTime, Session.getScriptTimeZone(), 'HH:mm') + 
                ' - ' + 
                Utilities.formatDate(event.endTime, Session.getScriptTimeZone(), 'HH:mm');
    }
    
    // สร้างเนื้อหาของ bubble
    var bodyContents = [
      {
        type: "text",
        text: event.title,
        weight: "bold",
        size: "lg",
        color: "#333333",
        wrap: true,
        margin: "sm"
      }
    ];
    
    // เพิ่มเวลา
    bodyContents.push({
      type: "box",
      layout: "horizontal",
      contents: [
        {
          type: "text",
          text: "⏰",
          size: "md",
          flex: 0,
          color: "#FF6B6B"
        },
        {
          type: "text",
          text: timeText,
          size: "md",
          color: "#555555",
          flex: 1,
          weight: "bold",
          margin: "sm"
        }
      ],
      margin: "md"
    });
    
    // เพิ่มสถานที่ (ถ้ามี)
    if (event.location) {
      bodyContents.push({
        type: "box",
        layout: "horizontal",
        contents: [
          {
            type: "text",
            text: "📍",
            size: "md",
            flex: 0,
            color: "#4ECDC4"
          },
          {
            type: "text",
            text: event.location,
            size: "sm",
            color: "#666666",
            flex: 1,
            wrap: true,
            margin: "sm"
          }
        ],
        margin: "md"
      });
    }
    
    // เพิ่มรายละเอียด (ถ้ามี)
    if (event.description) {
      // ตัดข้อความให้สั้นลง
      var description = event.description;
      if (description.length > 80) {
        description = description.substring(0, 80) + "...";
      }
      
      bodyContents.push({
        type: "box",
        layout: "horizontal",
        contents: [
          {
            type: "text",
            text: "📝",
            size: "md",
            flex: 0,
            color: "#45B7D1"
          },
          {
            type: "text",
            text: description,
            size: "sm",
            color: "#666666",
            flex: 1,
            wrap: true,
            margin: "sm"
          }
        ],
        margin: "md"
      });
    }
    
    // เพิ่มผู้สร้าง (ถ้ามี)
    if (event.creator) {
      bodyContents.push({
        type: "box",
        layout: "horizontal",
        contents: [
          {
            type: "text",
            text: "👤",
            size: "md",
            flex: 0,
            color: "#A29BFE"
          },
          {
            type: "text",
            text: event.creator,
            size: "sm",
            color: "#666666",
            flex: 1,
            margin: "sm"
          }
        ],
        margin: "md"
      });
    }
    
    // เพิ่มผู้เข้าร่วม (ถ้ามี)
    if (event.guests && event.guests.length > 0) {
      var guestText = event.guests.slice(0, 2).join(", ");
      if (event.guests.length > 2) {
        guestText += " และอีก " + (event.guests.length - 2) + " คน";
      }
      
      bodyContents.push({
        type: "box",
        layout: "horizontal",
        contents: [
          {
            type: "text",
            text: "👥",
            size: "md",
            flex: 0,
            color: "#96CEB4"
          },
          {
            type: "text",
            text: guestText,
            size: "sm",
            color: "#666666",
            flex: 1,
            wrap: true,
            margin: "sm"
          }
        ],
        margin: "md"
      });
    }
    
    // เพิ่มข้อมูลวันที่สร้าง
    bodyContents.push({
      type: "separator",
      margin: "lg"
    });
    
    bodyContents.push({
      type: "text",
      text: "วันที่สร้าง: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      size: "xs",
      color: "#999999",
      align: "center",
      margin: "md"
    });
    
    var bubble = {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "⚠️ แจ้งเตือนการนัดหมาย",
            weight: "bold",
            size: "md",
            color: "#FFFFFF"
          },
          {
            type: "text",
            text: "📅 " + displayDate,
            weight: "bold",
            size: "sm",
            color: "#FFFFFF",
            margin: "xs"
          }
        ],
        backgroundColor: headerColor,
        paddingAll: "lg"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: bodyContents,
        backgroundColor: "#FFFFFF",
        paddingAll: "lg"
      }
    };
    
    bubbles.push(bubble);
  }
  
  // ถ้ามีกิจกรรมมากกว่า 1 ใช้ carousel
  if (bubbles.length > 1) {
    return {
      type: "flex",
      altText: "รายละเอียดกิจกรรมวันที่ " + displayDate,
      contents: {
        type: "carousel",
        contents: bubbles
      }
    };
  } else {
    return {
      type: "flex",
      altText: "รายละเอียดกิจกรรมวันที่ " + displayDate,
      contents: bubbles[0]
    };
  }
}

/**
 * กำหนดสีหัวข้อตามประเภทกิจกรรม (สีเข้มที่ตัดกันชัด)
 */
function getEventHeaderColor(title) {
  var lowerTitle = title.toLowerCase();
  
  // ประชุม
  if (lowerTitle.includes('ประชุม') || lowerTitle.includes('meeting')) {
    return "#E74C3C"; // สีแดงเข้ม
  }
  // งานเยี่ยม/เยือน
  else if (lowerTitle.includes('เยี่ยม') || lowerTitle.includes('เยือน') || lowerTitle.includes('visit')) {
    return "#1ABC9C"; // สีเขียวมิ้นท์เข้ม
  }
  // อบรม/ฝึกอบรม
  else if (lowerTitle.includes('อบรม') || lowerTitle.includes('ฝึก') || lowerTitle.includes('training')) {
    return "#3498DB"; // สีฟ้าเข้ม
  }
  // วันหยุด/ลา
  else if (lowerTitle.includes('หยุด') || lowerTitle.includes('ลา') || lowerTitle.includes('holiday')) {
    return "#27AE60"; // สีเขียวเข้ม
  }
  // งานวันเกิด/ปาร์ตี้
  else if (lowerTitle.includes('วันเกิด') || lowerTitle.includes('ปาร์ตี้') || lowerTitle.includes('party')) {
    return "#F39C12"; // สีส้มเข้ม
  }
  // งานนำเสนอ
  else if (lowerTitle.includes('นำเสนอ') || lowerTitle.includes('presentation')) {
    return "#9B59B6"; // สีม่วงเข้ม
  }
  // ค่าเริ่มต้น
  else {
    return "#34495E"; // สีเทาน้ำเงินเข้ม
  }
}

/**
 * สร้าง Flex Message แสดงกิจกรรมทั้งเดือน
 */
function createMonthlyEventsFlexMessage(month, year) {
  var monthNames = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  
  // ดึงข้อมูลกิจกรรมทั้งเดือน
  var monthlyEvents = getCalendarEventsForMonth(year, month);
  var allEvents = getDetailedMonthlyEvents(year, month);
  
  if (allEvents.length === 0) {
    return {
      type: 'text',
      text: '📅 ' + monthNames[month - 1] + ' ' + (year + 543) + '\n\n❌ ไม่มีกิจกรรมในเดือนนี้'
    };
  }
  
  // จัดกลุ่มกิจกรรมตามวันที่
  var eventsByDate = {};
  for (var i = 0; i < allEvents.length; i++) {
    var event = allEvents[i];
    var day = event.startTime.getDate();
    
    if (!eventsByDate[day]) {
      eventsByDate[day] = [];
    }
    eventsByDate[day].push(event);
  }
  
  var bubbles = [];
  var sortedDays = Object.keys(eventsByDate).sort(function(a, b) { return parseInt(a) - parseInt(b); });
  
  // สร้าง bubble สำหรับแต่ละวัน (แสดงสูงสุด 10 วัน)
  for (var i = 0; i < Math.min(sortedDays.length, 10); i++) {
    var day = sortedDays[i];
    var dayEvents = eventsByDate[day];
    var displayDate = day + '/' + month + '/' + year;
    
    // สร้างเนื้อหาสำหรับวันนี้
    var contents = [
      {
        type: "text",
        text: "📅 วันที่ " + day,
        weight: "bold",
        size: "lg",
        color: "#333333"
      }
    ];
    
    // เพิ่มกิจกรรมของวันนี้
    for (var j = 0; j < Math.min(dayEvents.length, 5); j++) { // แสดงสูงสุด 5 กิจกรรมต่อวัน
      var event = dayEvents[j];
      
      var timeText = "";
      if (event.isAllDay) {
        timeText = "ตลอดวัน";
      } else {
        timeText = Utilities.formatDate(event.startTime, Session.getScriptTimeZone(), 'HH:mm') + 
                  '-' + 
                  Utilities.formatDate(event.endTime, Session.getScriptTimeZone(), 'HH:mm');
      }
      
      // เพิ่มกิจกรรม
      contents.push({
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: event.title,
            weight: "bold",
            size: "sm",
            color: "#555555",
            wrap: true
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              {
                type: "text",
                text: "⏰",
                size: "sm",
                flex: 0,
                color: "#FF6B6B"
              },
              {
                type: "text",
                text: timeText,
                size: "sm",
                color: "#666666",
                flex: 1,
                margin: "sm"
              }
            ]
          }
        ],
        margin: "md",
        paddingAll: "sm",
        backgroundColor: "#F8F9FA",
        cornerRadius: "md"
      });
    }
    
    // ถ้ามีกิจกรรมเกิน 5 แสดงจำนวนที่เหลือ
    if (dayEvents.length > 5) {
      contents.push({
        type: "text",
        text: "และอีก " + (dayEvents.length - 5) + " กิจกรรม...",
        size: "xs",
        color: "#999999",
        align: "center",
        margin: "sm"
      });
    }
    
    var bubble = {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: monthNames[month - 1] + " " + (year + 543),
            weight: "bold",
            size: "md",
            color: "#FFFFFF",
            align: "center"
          }
        ],
        backgroundColor: "#34495E",
        paddingAll: "lg"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: contents,
        paddingAll: "lg"
      },
      footer: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "button",
            action: {
              type: "postback",
              label: "ดูรายละเอียดวันที่ " + day,
              data: "select_date_" + year + "_" + month + "_" + day
            },
            style: "primary",
            color: "#34495E"
          }
        ],
        paddingAll: "md"
      }
    };
    
    bubbles.push(bubble);
  }
  
  // ถ้ามีมากกว่า 10 วัน แสดงข้อความเพิ่มเติม
  if (sortedDays.length > 10) {
    var additionalBubble = {
      type: "bubble",
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "📋 รายการเพิ่มเติม",
            weight: "bold",
            size: "lg",
            color: "#333333",
            align: "center"
          },
          {
            type: "text",
            text: "มีกิจกรรมอีก " + (sortedDays.length - 10) + " วัน\nในเดือน" + monthNames[month - 1],
            size: "sm",
            color: "#666666",
            align: "center",
            wrap: true,
            margin: "md"
          }
        ],
        paddingAll: "lg"
      }
    };
    bubbles.push(additionalBubble);
  }
  
  return {
    type: "flex",
    altText: "กิจกรรมประจำเดือน " + monthNames[month - 1] + " " + (year + 543),
    contents: {
      type: "carousel",
      contents: bubbles
    }
  };
}

/**
 * ดึงข้อมูลกิจกรรมรายละเอียดทั้งเดือน
 */
function getDetailedMonthlyEvents(year, month) {
  try {
    var calendar = CalendarApp.getCalendarById(APPOINTMENT_CALENDAR_ID);
    var startDate = new Date(year, month - 1, 1);
    var endDate = new Date(year, month, 0, 23, 59, 59);
    
    var events = calendar.getEvents(startDate, endDate);
    var eventDetails = [];
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      eventDetails.push({
        title: event.getTitle(),
        description: event.getDescription() || '',
        location: event.getLocation() || '',
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        isAllDay: event.isAllDayEvent()
      });
    }
    
    return eventDetails;
  } catch (error) {
    console.error('Error getting detailed monthly events:', error);
    return [];
  }
}

/**
 * บันทึกลง Google Sheets
 */
function saveToGoogleSheets(date, userId) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    var timestamp = new Date();
    
    // เพิ่มข้อมูลในแถวใหม่
    sheet.appendRow([
      timestamp,           // วันเวลาที่บันทึก
      userId,             // LINE User ID
      date,               // วันที่ที่เลือก
      'จองวันที่'         // หมายเหตุ
    ]);
    
    console.log('บันทึกข้อมูลลง Google Sheets สำเร็จ: ' + date);
  } catch (error) {
    console.error('Error saving to Google Sheets:', error);
  }
}

/**
 * เพิ่มลงปฏิทิน Google Calendar (ตัวเลือก)
 */
function addToGoogleCalendar(date, title) {
  try {
    var calendar = CalendarApp.getCalendarById(APPOINTMENT_CALENDAR_ID);
    var eventDate = new Date(date);
    
    calendar.createEvent(
      title || 'เหตุการณ์จาก LINE OA',
      eventDate,
      eventDate,
      {
        description: 'จองผ่าน LINE OA Bot เมื่อ ' + new Date().toLocaleString('th-TH')
      }
    );
    
    console.log('เพิ่มลงปฏิทินสำเร็จ: ' + date);
  } catch (error) {
    console.error('Error adding to Google Calendar:', error);
  }
}

// /**
//  * ฟังก์ชันช่วยเหลือ - เพิ่ม 0 หน้าเลข
//  */
// function padZero(num) {
//   return num < 10 ? '0' + num : num.toString();
// }

// /**
//  * ส่งข้อความ Flex Message
//  */
// function replyFlexMessage(token, flexMessage) {
//   const replyData = {
//     replyToken: token,
//     messages: [flexMessage]
//   };

//   const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
//     headers: {
//       'Content-Type': 'application/json',
//       'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
//     },
//     method: 'POST',
//     payload: JSON.stringify(replyData),
//     muteHttpExceptions: true
//   });

//   if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
//     console.log('✅ Flex message sent successfully');
//     return ContentService.createTextOutput(JSON.stringify({'status': 'ok'})).setMimeType(ContentService.MimeType.JSON);
//   } else {
//     console.log('❌ Error sending flex message:', response.getContentText());
//     return ContentService.createTextOutput(JSON.stringify({'status': 'error'})).setMimeType(ContentService.MimeType.JSON);
//   }
// }

// ===== ฟังก์ชันช่วยทดสอบ =====

/**
 * ทดสอบระบบค้นหา USER ID
 */
function testSearchSystem() {
  try {
    console.log('🧪 Testing USER ID search system...');
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userIdSheet = ss.getSheetByName("USER ID");
    
    if (!userIdSheet) {
      return { success: false, error: 'USER ID sheet not found' };
    }
    
    const lastRow = userIdSheet.getLastRow();
    console.log('📊 Total rows:', lastRow);
    
    if (lastRow >= 2) {
      const data = userIdSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      const validUsers = data.filter(row => 
        row[1] && row[1].toString().trim() !== "" && row[1].toString() !== "Group"
      );
      
      console.log('👥 Valid users found:', validUsers.length);
      
      // ทดสอบการค้นหา
      if (validUsers.length > 0) {
        const testName = "กมลภพ";
        const matches = findMatchingUsers(data, testName);
        
        console.log('🔍 Test search for "กมลภพ":', matches.length, 'matches found');
        matches.forEach((match, index) => {
          console.log(`  ${index + 1}. ${match.fullName} (${match.matchType})`);
        });
      }
      
      return { 
        success: true, 
        totalRows: lastRow,
        validUsers: validUsers.length,
        sampleUsers: validUsers.slice(0, 5).map(row => row[1])
      };
    } else {
      return { success: false, error: 'No data found' };
    }
    
  } catch (error) {
    console.error('❌ Error in testSearchSystem:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * สร้างกิจกรรมทดสอบในปฏิทิน
 */
function createTestEvents() {
  try {
    var calendar = CalendarApp.getCalendarById(APPOINTMENT_CALENDAR_ID);
    var today = new Date();
    
    // สร้างกิจกรรมทดสอบ
    var event1 = calendar.createEvent(
      'ประชุมทีมพัฒนา',
      new Date(today.getTime() + (1 * 24 * 60 * 60 * 1000)), // พรุ่งนี้
      new Date(today.getTime() + (1 * 24 * 60 * 60 * 1000) + (2 * 60 * 60 * 1000)), // 2 ชั่วโมง
      {
        description: 'ประชุมร่วมกันเพื่อหารือเกี่ยวกับการพัฒนาระบบใหม่ รวมถึงการวางแผนงานในไตรมาสต่อไป',
        location: 'ห้องประชุม A ชั้น 3',
        guests: 'developer1@company.com,developer2@company.com'
      }
    );
    
    var event2 = calendar.createEvent(
      'การเยี่ยมลูกค้า ABC Corp',
      new Date(today.getTime() + (2 * 24 * 60 * 60 * 1000)), // มะรืนนี้
      new Date(today.getTime() + (2 * 24 * 60 * 60 * 1000) + (3 * 60 * 60 * 1000)), // 3 ชั่วโมง
      {
        description: 'นำเสนอโซลูชันใหม่และติดตามความคืบหน้าโครงการ พร้อมทั้งหารือแผนการขยายธุรกิจ',
        location: 'สำนักงาน ABC Corp ชั้น 15 อาคาร XYZ',
        guests: 'sales@company.com,manager@company.com'
      }
    );
    
    var event3 = calendar.createEvent(
      'อบรมเทคโนโลยีใหม่',
      new Date(today.getTime() + (3 * 24 * 60 * 60 * 1000)), // วันที่ 3
      new Date(today.getTime() + (3 * 24 * 60 * 60 * 1000) + (6 * 60 * 60 * 1000)), // 6 ชั่วโมง
      {
        description: 'อบรมการใช้งาน AI และ Machine Learning ในการพัฒนาธุรกิจ',
        location: 'โรงแรม Grand Palace ห้องประชุมใหญ่',
        guests: 'hr@company.com,all-staff@company.com'
      }
    );
    
    console.log('Created test events successfully');
    console.log('Event 1 ID:', event1.getId());
    console.log('Event 2 ID:', event2.getId());
    console.log('Event 3 ID:', event3.getId());
    
    return { 
      success: true, 
      events: [
        { id: event1.getId(), title: 'ประชุมทีมพัฒนา' },
        { id: event2.getId(), title: 'การเยี่ยมลูกค้า ABC Corp' },
        { id: event3.getId(), title: 'อบรมเทคโนโลยีใหม่' }
      ]
    };
    
  } catch (error) {
    console.error('Error creating test events:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ทดสอบการทำงานของปฏิทิน
 */
function testCalendar() {
  try {
    var today = new Date();
    var calendar = createCalendarFlexMessage(today.getFullYear(), today.getMonth() + 1);
    console.log('📅 Calendar created successfully');
    console.log(JSON.stringify(calendar, null, 2));
    return { success: true, calendar: calendar };
  } catch (error) {
    console.error('Error testing calendar:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ทดสอบการดึงข้อมูลจาก Google Calendar
 */
function testCalendarEvents() {
  try {
    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth() + 1;
    
    console.log('Testing calendar events for:', year, month);
    var events = getCalendarEventsForMonth(year, month);
    console.log('Events found:', JSON.stringify(events, null, 2));
    
    // ทดสอบดึงข้อมูลวันที่เฉพาะ
    var dateString = year + '-' + padZero(month) + '-' + padZero(today.getDate());
    console.log('Testing events for date:', dateString);
    var dayEvents = getDetailedCalendarEventsForDate(dateString);
    console.log('Day events:', JSON.stringify(dayEvents, null, 2));
    
    // ทดสอบสร้าง Flex Message
    if (dayEvents.length > 0) {
      var displayDate = today.getDate() + '/' + month + '/' + year;
      var flexMessage = createEventDetailsFlexMessage(displayDate, dayEvents);
      console.log('Flex Message created successfully');
    }
    
    return { 
      success: true, 
      monthlyEvents: events,
      dayEvents: dayEvents,
      testDate: dateString
    };
    
  } catch (error) {
    console.error('Error testing calendar events:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ทดสอบการแสดงกิจกรรมทั้งเดือน
 */
function testMonthlyEvents() {
  try {
    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth() + 1;
    
    console.log('Testing monthly events for:', month, year);
    var monthlyMessage = createMonthlyEventsFlexMessage(month, year);
    console.log('Monthly Events Message created successfully');
    
    return { 
      success: true, 
      message: monthlyMessage
    };
    
  } catch (error) {
    console.error('Error testing monthly events:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ทดสอบรูปแบบวันที่
 */
function testDateFormats() {
  var testDates = [
    '12/7/2568',
    '12 กรกฎาคม 2568', 
    '7/2568',
    'กรกฎาคม 2568',
    '25-12-2568',
    '1 มกราคม 2569',
    'invalid date'
  ];
  
  console.log('Testing date formats:');
  var results = [];
  
  for (var i = 0; i < testDates.length; i++) {
    var testDate = testDates[i];
    var isValid = isDateFormatForCalendar(testDate);
    var parsed = parseThaiDateForCalendar(testDate);
    
    var result = {
      input: testDate,
      isValid: isValid,
      parsed: parsed
    };
    
    results.push(result);
    console.log(testDate + ':', 'isValid=' + isValid + ', parsed=' + JSON.stringify(parsed));
  }
  
/**
 * ส่งข้อความตอบกลับธรรมดา
 */
function replyMessage2(token, message) {
  try {
    const replyData = {
      replyToken: token,
      messages: [typeof message === 'string' ? { type: 'text', text: message } : message]
    };

    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
      },
      method: 'POST',
      payload: JSON.stringify(replyData),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      console.log('✅ Message sent successfully');
      return ContentService.createTextOutput(JSON.stringify({'status': 'ok'})).setMimeType(ContentService.MimeType.JSON);
    } else {
      console.log('❌ Error sending message:', response.getContentText());
      return ContentService.createTextOutput(JSON.stringify({'status': 'error'})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    console.error('❌ Error in replyMessage:', error);
    return ContentService.createTextOutput(JSON.stringify({'status': 'error'})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ดึงข้อมูลโปรไฟล์ผู้ใช้หลายคน (รูปแบบเก่า)
 */
function getUserProfiles(userId) {
  try {
    const profile = getUserProfile(userId);
    if (profile) {
      return [profile.displayName, profile.pictureUrl];
    }
    return ['Unknown User', ''];
  } catch (error) {
    console.error('Error getting user profiles:', error);
    return ['Unknown User', ''];
  }
}

/**
 * จัดรูปแบบวันที่เป็นภาษาไทย
 */
function formatThaiDate(date) {
  try {
    const thaiMonths = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    
    const day = date.getDate();
    const month = thaiMonths[date.getMonth()];
    const year = date.getFullYear() + 543;
    
    return `${day} ${month} ${year}`;
  } catch (error) {
    console.error('Error formatting Thai date:', error);
    return date.toString();
  }
}

// ===== ฟังก์ชันเพิ่มเติม =====

/**
 * เพิ่มข้อมูลตัวอย่างสำหรับทดสอบ
 */
function addTestUsers() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userIdSheet = ss.getSheetByName("USER ID");
    
    const sampleUsers = [
      ['U1111111111111111', 'กมลภพ จำปา'],
      ['U2222222222222222', 'กมลภพ จำปี'],
      ['U3333333333333333', 'สมชาย ใจดี'],
      ['U4444444444444444', 'สมหญิง รักดี'],
      ['U5555555555555555', 'ประยุท สุขใส']
    ];
    
    const lastRow = userIdSheet.getLastRow();
    const nextRow = lastRow + 1;
    
    userIdSheet.getRange(nextRow, 1, sampleUsers.length, 2).setValues(sampleUsers);
    
    console.log('✅ Added', sampleUsers.length, 'test users');
    return { success: true, addedUsers: sampleUsers.length };
    
  } catch (error) {
    console.error('❌ Error in addTestUsers:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ลบข้อมูลตัวอย่าง
 */
function removeTestUsers() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userIdSheet = ss.getSheetByName("USER ID");
    
    const lastRow = userIdSheet.getLastRow();
    const data = userIdSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    let removedCount = 0;
    
    // ลบแถวที่มี USER ID ขึ้นต้นด้วย "U111111", "U222222" เป็นต้น
    for (let i = data.length - 1; i >= 0; i--) {
      const userId = data[i][0] ? data[i][0].toString() : '';
      if (userId.match(/^U[1-5]{15}$/)) {
        userIdSheet.deleteRow(i + 2);
        removedCount++;
      }
    }
    
    console.log('✅ Removed', removedCount, 'test users');
    return { success: true, removedCount: removedCount };
    
  } catch (error) {
    console.error('❌ Error in removeTestUsers:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ทดสอบการทำงานครบวงจร
 */
function runCompleteTest() {
  try {
    console.log('🧪 Starting complete system test...');
    
    // 1. ทดสอบระบบค้นหา USER ID
    console.log('📋 Testing USER ID search system...');
    const searchTest = testSearchSystem();
    
    // 2. ทดสอบระบบปฏิทิน
    console.log('📅 Testing calendar system...');
    const calendarTest = testCalendar();
    
    // 3. ทดสอบการดึงข้อมูลจาก Google Calendar
    console.log('📊 Testing calendar events...');
    const eventsTest = testCalendarEvents();
    
    // 4. ทดสอบรูปแบบวันที่
    console.log('📝 Testing date formats...');
    const dateTest = testDateFormats();
    
    console.log('✅ Complete test finished');
    
    return {
      success: true,
      results: {
        userSearch: searchTest,
        calendar: calendarTest,
        events: eventsTest,
        dateFormats: dateTest
      }
    };
    
  } catch (error) {
    console.error('❌ Error in complete test:', error);
    return { success: false, error: error.toString() };
  }
}
}

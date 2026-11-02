/**
 * โค้ด Google Apps Script สำหรับระบบเช็กชื่อผ่านใบหน้า (Face Recognition Attendance)
 * วิชา: เทคโนโลยีกลุ่มเมฆ
 * 
 * วิธีการนำไปใช้:
 * 1. เปิด Google Sheets ใหม่ (https://sheets.new)
 * 2. ไปที่เมนู "ส่วนขยาย" (Extensions) > "Apps Script"
 * 3. ลบโค้ดเดิมทั้งหมดออก แล้ววางโค้ดชุดนี้ลงไปแทน
 * 4. กดปุ่ม "บันทึก" (รูปแผ่นดิสก์)
 * 5. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) > "การทำให้ใช้งานได้ใหม่" (New deployment)
 * 6. เลือกประเภทเป็น "เว็บแอป" (Web app)
 *    - คำอธิบาย: ระบบเช็กชื่อ
 *    - ปฏิบัติการในฐานะ: ตัวฉัน (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน (Anyone)  <--- สำคัญมาก ต้องเลือก "ทุกคน"
 * 7. กด "ทำให้ใช้งานได้" (Deploy) แล้วคัดลอก "URL ของเว็บแอป" ไปใส่ในหน้าเว็บ index.html
 */

function doPost(e) {
  return handleAttendanceData(e);
}

function doGet(e) {
  // รองรับการส่งข้อมูลผ่าน GET หรือใช้ทดสอบการเชื่อมต่อ
  if (e && e.parameter && e.parameter.name) {
    return handleAttendanceData(e);
  }
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    message: "Google Apps Script สำหรับระบบเช็กชื่อพร้อมใช้งานแล้ว!"
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleAttendanceData(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = {};

    // อ่านข้อมูลจาก POST Body (JSON) หรือจาก URL Parameter
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e.parameter) {
      data = e.parameter;
    }

    // ถ้ายังเป็นตารางว่าง ให้สร้างหัวตาราง (Header) สวยงาม
    if (sheet.getLastRow() === 0) {
      var headers = ["ลำดับ", "วันที่", "เวลา", "ชื่อผู้เข้าเรียน", "ความมั่นใจ AI (%)", "สถานะ", "บันทึกระบบเมื่อ"];
      sheet.appendRow(headers);
      
      // ตกแต่งหัวตาราง
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#2563EB"); // สีน้ำเงิน
      headerRange.setFontColor("#FFFFFF");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // ลำดับที่ของแถวใหม่
    var rowNumber = sheet.getLastRow();
    var now = new Date();
    
    // จัดรูปแบบวันที่และเวลาประเทศไทย (Asia/Bangkok)
    var dateStr = data.date || Utilities.formatDate(now, "Asia/Bangkok", "dd/MM/yyyy");
    var timeStr = data.time || Utilities.formatDate(now, "Asia/Bangkok", "HH:mm:ss");
    var name = data.name || "ไม่ระบุชื่อ";
    var confidence = data.confidence ? data.confidence + "%" : "-";
    var status = data.status || "มาเรียน";

    // บันทึกแถวข้อมูลลง Google Sheet
    sheet.appendRow([
      rowNumber,
      dateStr,
      timeStr,
      name,
      confidence,
      status,
      now
    ]);

    // จัดข้อความให้อยู่กึ่งกลางในช่องลำดับ วันที่ เวลา สถานะ
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setHorizontalAlignment("center");
    sheet.getRange(lastRow, 2).setHorizontalAlignment("center");
    sheet.getRange(lastRow, 3).setHorizontalAlignment("center");
    sheet.getRange(lastRow, 5).setHorizontalAlignment("center");
    sheet.getRange(lastRow, 6).setHorizontalAlignment("center");

    // ส่งข้อความผลลัพธ์กลับไปยังหน้าเว็บ
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "บันทึกการเช็กชื่อสำเร็จเรียบร้อย",
      row: rowNumber,
      name: name,
      time: timeStr
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "เกิดข้อผิดพลาด: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

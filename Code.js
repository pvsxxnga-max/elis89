// ==========================================
// e-LIS: Backend Services (Google Apps Script)
// ==========================================

function doGet(e) {
  // ตรวจสอบการกดปุ่ม "ล็อกการใช้งาน" จากอีเมล
  if (e.parameter.action === 'lock' && e.parameter.email) {
    lockUserAccount(e.parameter.email);
    return HtmlService.createHtmlOutput('<div style="text-align:center; padding:50px; font-family:sans-serif;"><h2>บัญชีถูกล็อกเรียบร้อยแล้ว</h2><p>กรุณาติดต่อผู้ดูแลระบบเพื่อปลดล็อก</p></div>');
  }

  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('e-LIS ระบบสารสนเทศเพื่อการบริหารสินเชื่อ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------
// Email & Security Services
// ------------------------------------------

function sendLoginAlert(email, ip, role) {
  try {
    const lockUrl = ScriptApp.getService().getUrl() + "?action=lock&email=" + encodeURIComponent(email);
    const htmlBody = `
      <div style="font-family: 'Sarabun', sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 10px;">
        <h2 style="color: #1e3a8a; text-align: center;">e-LIS Security Alert</h2>
        <p>เรียนผู้ใช้งาน,</p>
        <p>ระบบตรวจพบการเข้าสู่ระบบใหม่ด้วยบัญชีของคุณ (${role})</p>
        <div style="background-color: #f8fafc; padding: 15px; border-radius: 8px; margin: 20px 0;">
          <p style="margin: 5px 0;"><strong>อีเมล:</strong> ${email}</p>
          <p style="margin: 5px 0;"><strong>IP Address:</strong> ${ip || 'ไม่สามารถระบุได้'}</p>
          <p style="margin: 5px 0;"><strong>เวลา:</strong> ${new Date().toLocaleString('th-TH')}</p>
        </div>
        <p style="color: #ef4444; font-weight: bold;">หากคุณไม่ได้เป็นผู้เข้าสู่ระบบ โปรดคลิกปุ่มด้านล่างทันทีเพื่อระงับการใช้งาน</p>
        <div style="text-align: center; margin-top: 30px;">
          <a href="${lockUrl}" style="background-color: #dc2626; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">ล็อกการใช้งานบัญชีนี้</a>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: "แจ้งเตือนการเข้าสู่ระบบใหม่ - e-LIS",
      htmlBody: htmlBody
    });
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function sendOTP(email) {
  try {
    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const htmlBody = `
      <div style="font-family: 'Sarabun', sans-serif; text-align: center; padding: 30px;">
        <h2>รหัสยืนยัน OTP ของคุณคือ</h2>
        <h1 style="color: #1e3a8a; font-size: 40px; letter-spacing: 5px;">${otp}</h1>
        <p>รหัสนี้ใช้สำหรับยืนยันตัวตนในระบบ e-LIS กรุณาอย่าเปิดเผยให้ผู้อื่นทราบ</p>
      </div>
    `;
    MailApp.sendEmail({
      to: email,
      subject: "รหัส OTP สำหรับระบบ e-LIS",
      htmlBody: htmlBody
    });
    return { success: true, otp: otp }; // In production, store OTP in Firebase/Properties, don't return to client. Returning for standalone completeness.
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function sendStatusUpdateEmail(email, loanId, statusText, note) {
  try {
    const htmlBody = `
      <div style="font-family: 'Sarabun', sans-serif; padding: 20px;">
        <h2 style="color: #1e3a8a;">อัปเดตสถานะการยื่นสินเชื่อ</h2>
        <p>เลขที่คำขอ: <strong>${loanId}</strong></p>
        <p>สถานะปัจจุบัน: <strong style="color: #059669;">${statusText}</strong></p>
        ${note ? `<p>หมายเหตุจากเจ้าหน้าที่: <em>${note}</em></p>` : ''}
        <p>คุณสามารถตรวจสอบรายละเอียดเพิ่มเติมได้ที่ระบบ e-LIS</p>
      </div>
    `;
    MailApp.sendEmail({
      to: email,
      subject: `อัปเดตสถานะคำขอสินเชื่อ ${loanId}`,
      htmlBody: htmlBody
    });
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function lockUserAccount(email) {
  // Logic สำหรับบันทึกสถานะ 'Locked' ลง Firebase
  // เนื่องจากฟังก์ชันนี้เรียกจาก Server-side, คุณสามารถใช้ UrlFetchApp ยิง REST API ไปยัง Firebase ได้
  PropertiesService.getScriptProperties().setProperty(`LOCKED_${email}`, 'true');
}

// ------------------------------------------
// Export Services (PDF/Word)
// ------------------------------------------
function exportReport(type, dataStr) {
  // สร้างรายงาน PDF แบบจำลอง (เนื่องจาก GAS สร้าง PDF ต้องใช้ Google Docs เป็น Template)
  // ในที่นี้จะสร้างไฟล์และคืนค่า URL ของไฟล์
  const data = JSON.parse(dataStr);
  const doc = DocumentApp.create('e-LIS_Report_' + new Date().getTime());
  const body = doc.getBody();
  
  body.insertParagraph(0, 'รายงานสรุปสถานะสินเชื่อ (e-LIS)').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`วันที่พิมพ์: ${new Date().toLocaleString('th-TH')}`);
  
  data.forEach(item => {
    body.appendParagraph(`เลขที่: ${item.id} | สินเชื่อ: ${item.name} | สถานะ: ${item.status}`);
  });
  
  doc.saveAndClose();
  
  if (type === 'pdf') {
    const pdf = doc.getAs('application/pdf');
    const file = DriveApp.createFile(pdf);
    return file.getUrl();
  }
  
  return doc.getUrl();
}

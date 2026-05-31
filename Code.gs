/**
 * MCUKS One-Year Life Balance System
 * ระบบแผนชีวิต 1 ปี : ธรรมวินัยเป็นหลัก งานมีระบบ ใจมีสติ
 */

// 1. ตัวแปรตั้งค่าหลัก (ใช้เป็นค่าเริ่มต้น หากยังไม่มีในชีต Settings)
const CONFIG = {
  SHEET_ID: '1BL2XJ-UPwxrMz0gZFxmMkB1Cmw8foIUawGy2l97bkFI', // เปลี่ยนเป็น ID ของคุณ
  SCRIPT_ID: '1F1NrPf21Vww_aCA7y9Zj9zUZWz5Vz4xWRdKPJg5I-l0Qjxki-1ZsibLW',
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbxxtOYCPU5EGKhVmCKeMuQOjBJj8mkwO35H4T-tj6Qmmtqpt2Sh94af9NnsUE6eRqVwyA/exec',
  LOGO_URL: 'https://lh3.googleusercontent.com/d/1Oes1htLwaNV_qE7eJkuSlPMHqXqJJ9ec',
  OWNER_IMAGE_URL: 'https://lh3.googleusercontent.com/d/1Oes1htLwaNV_qE7eJkuSlPMHqXqJJ9ec',
  ORGANIZATION_NAME: 'หน่วยวิทยบริการ วิทยาลัยสงฆ์ขอนแก่น จังหวัดกาฬสินธุ์',
  UNIVERSITY_NAME: 'มหาวิทยาลัยมหาจุฬาลงกรณราชวิทยาลัย',
  SYSTEM_NAME: 'ระบบแผนชีวิต 1 ปี : ธรรมวินัยเป็นหลัก งานมีระบบ ใจมีสติ',
  USER_NAME: 'พระมหาธงชัย วิลาสินี',
  USER_ROLE: 'นักวิชาการศึกษา / เจ้าหน้าที่สารบรรณ ทะเบียนและวัดผล',
  TEMPLE_NAME: 'วัดกลาง พระอารามหลวง จ.กาฬสินธุ์',
  START_DATE: '2026-06-01',
  END_DATE: '2027-05-31',
  MORNING_CHANT_TIME: '08:00',
  EVENING_CHANT_TIME: '18:00',
  WORK_DAYS: 'Monday,Tuesday,Friday,Saturday,Sunday',
  OFF_DAYS: 'Wednesday,Thursday',
  TIMEZONE: 'Asia/Bangkok'
};

// 2. ฟังก์ชันหลักสำหรับ Web App
function doGet(e) {
  let template = HtmlService.createTemplateFromFile('index');
  // ส่งค่าเริ่มต้นไปที่หน้าเว็บ
  template.config = getSettings(); 
  return template.evaluate()
    .setTitle('ระบบแผนชีวิต 1 ปี')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 3. ฟังก์ชันดึง/บันทึกการตั้งค่า
function getSettings() {
  try {
    let sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('Settings');
    if (!sheet) return CONFIG;
    let data = sheet.getDataRange().getValues();
    let settings = {};
    for (let i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }
    // ใช้ค่าจาก CONFIG เป็น Fallback ถ้าในชีตไม่มี
    return Object.keys(settings).length > 0 ? { ...CONFIG, ...settings } : CONFIG;
  } catch (e) {
    return CONFIG;
  }
}

function saveSettings(formObject) {
  let ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName('Settings');
  sheet.clear();
  sheet.appendRow(['SettingName', 'SettingValue', 'Description']);
  for (let key in formObject) {
    sheet.appendRow([key, formObject[key], '']);
  }
  return { status: 'success', message: 'บันทึกการตั้งค่าสำเร็จ' };
}

// 4. ฟังก์ชันสร้างฐานข้อมูลเริ่มต้น (Setup Sheets)
function setupSheets() {
  let ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheetsDef = [
    { name: 'Settings', headers: ['SettingName', 'SettingValue', 'Description'] },
    { name: 'Calendar', headers: ['Date', 'DayName', 'ThaiDate', 'IsWorkday', 'IsWanPhra', 'LunarText', 'SpecialDay'] },
    { name: 'WanPhraCalendar', headers: ['Date', 'ThaiDate', 'DayNameTH', 'YearBE', 'LunarText', 'SpecialDay'] },
    { name: 'DailyChecklist', headers: ['Timestamp', 'Date', 'Category', 'Item', 'IsDone'] },
    { name: 'HealthLog', headers: ['Timestamp', 'Date', 'Weight', 'Waist', 'SleepHours', 'SleepQuality', 'WalkMinutes', 'SweetDrink', 'WaterEnough', 'BodyStatus'] },
    { name: 'IFLog', headers: ['Timestamp', 'Date', 'IFType', 'EatingStart', 'EatingEnd', 'FastingSuccess', 'Symptom'] },
    { name: 'FinanceLog', headers: ['Timestamp', 'Date', 'Item', 'Amount', 'Category', 'Necessity', 'Note'] },
    { name: 'WorkLog', headers: ['Timestamp', 'Date', 'MainTask', 'CompletedTask', 'PendingTask', 'AIUsed', 'WorkCategory'] },
    { name: 'MindLog', headers: ['Timestamp', 'Date', 'GoodThing', 'TiredThing', 'LetGoThing', 'Gratitude', 'KindnessAction', 'TomorrowFocus'] },
    { name: 'RiskAlert', headers: ['Timestamp', 'Date', 'AlertType', 'Detail', 'Recommendation', 'Resolved'] }
  ];

  sheetsDef.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.appendRow(def.headers);
      sheet.getRange(1, 1, 1, def.headers.length).setFontWeight('bold').setBackground('#f3e5f5');
    }
  });
  
  // ลบชีตเริ่มต้น (Sheet1) ถ้ามี
  let sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1) ss.deleteSheet(sheet1);
  
  return "สร้างโครงสร้างชีตสำเร็จ!";
}

// 5. Seed ข้อมูลวันพระ 2569-2570
function seedWanPhraCalendar() {
  let sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('WanPhraCalendar');
  // ข้อมูลวันพระจาก Prompt
  const wanPhraData = [
    ['2026-06-08', '8 มิถุนายน 2569', 'จันทร์', 2569, 'แรม ๘ ค่ำ เดือนเจ็ด (๗)', 'วันอัฏฐมีบูชา'],
    ['2026-06-14', '14 มิถุนายน 2569', 'อาทิตย์', 2569, 'แรม ๑๔ ค่ำ เดือนเจ็ด (๗)', ''],
    ['2026-06-22', '22 มิถุนายน 2569', 'จันทร์', 2569, 'ขึ้น ๘ ค่ำ เดือนแปด (๘)', ''],
    ['2026-06-29', '29 มิถุนายน 2569', 'จันทร์', 2569, 'ขึ้น ๑๕ ค่ำ เดือนแปด (๘)', ''],
    ['2026-07-07', '7 กรกฎาคม 2569', 'อังคาร', 2569, 'แรม ๘ ค่ำ เดือนแปด (๘)', ''],
    ['2026-07-14', '14 กรกฎาคม 2569', 'อังคาร', 2569, 'แรม ๑๕ ค่ำ เดือนแปด (๘)', ''],
    ['2026-07-22', '22 กรกฎาคม 2569', 'พุธ', 2569, 'ขึ้น ๘ ค่ำ เดือนแปดหลัง (๘๘)', ''],
    ['2026-07-29', '29 กรกฎาคม 2569', 'พุธ', 2569, 'ขึ้น ๑๕ ค่ำ เดือนแปดหลัง (๘๘)', 'วันอาสาฬหบูชา'],
    ['2026-07-30', '30 กรกฎาคม 2569', 'พฤหัสบดี', 2569, 'แรม ๑ ค่ำ เดือนแปดหลัง (๘๘)', 'วันเข้าพรรษา'],
    ['2027-02-21', '21 กุมภาพันธ์ 2570', 'อาทิตย์', 2570, 'ขึ้น ๑๕ ค่ำ เดือนสาม (๓)', 'วันมาฆบูชา'],
    ['2027-05-20', '20 พฤษภาคม 2570', 'พฤหัสบดี', 2570, 'ขึ้น ๑๕ ค่ำ เดือนหก (๖)', 'วันวิสาขบูชา']
    // เพิ่มเติมให้ครบตามต้องการในภายหลังผ่านระบบจัดการได้
  ];
  
  if (sheet.getLastRow() <= 1) {
    wanPhraData.forEach(row => sheet.appendRow(row));
  }
  return "เพิ่มข้อมูลวันพระเบื้องต้นสำเร็จ";
}

function seedInitialData() {
  setupSheets();
  seedWanPhraCalendar();
  // รันตั้งค่า Config ลงชีตครั้งแรก
  saveSettings(CONFIG);
  return "ระบบพร้อมใช้งาน!";
}

// 6. ดึงข้อมูล Dashboard ของวันนี้
function getTodayDashboardData() {
  let today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
  let sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('WanPhraCalendar');
  let data = sheet.getDataRange().getValues();
  
  let isWanPhra = false;
  let wanPhraText = "";
  
  for(let i=1; i<data.length; i++){
    let dateStr = "";
    if(data[i][0] instanceof Date) {
       dateStr = Utilities.formatDate(data[i][0], CONFIG.TIMEZONE, "yyyy-MM-dd");
    } else {
       dateStr = data[i][0];
    }
    
    if(dateStr === today) {
      isWanPhra = true;
      wanPhraText = data[i][4] + (data[i][5] ? " (" + data[i][5] + ")" : "");
      break;
    }
  }

  // หา Checklist ของวันนี้
  let chkSheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('DailyChecklist');
  let chkData = chkSheet.getDataRange().getValues();
  let doneCount = 0;
  let totalCount = 20; // สมมติว่าฐานคือ 20 ข้อ
  
  for(let i=1; i<chkData.length; i++){
    let dateStr = chkData[i][1] instanceof Date ? Utilities.formatDate(chkData[i][1], CONFIG.TIMEZONE, "yyyy-MM-dd") : chkData[i][1];
    if(dateStr === today && chkData[i][4] === true) {
      doneCount++;
    }
  }

  let score = Math.round((doneCount / totalCount) * 100);
  if(score > 100) score = 100;

  return {
    date: today,
    isWanPhra: isWanPhra,
    wanPhraText: wanPhraText,
    score: score,
    settings: getSettings()
  };
}

// 7. ระบบบันทึกข้อมูล (CRUD)
function saveChecklistData(date, items) {
  let ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName('DailyChecklist');
  let timestamp = new Date();
  
  // ลบของเก่าของวันนี้ก่อน (เพื่อให้เป็นการอัปเดต)
  let data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    let d = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], CONFIG.TIMEZONE, "yyyy-MM-dd") : data[i][1];
    if (d === date) {
      sheet.deleteRow(i + 1);
    }
  }

  // บันทึกใหม่
  items.forEach(item => {
    sheet.appendRow([timestamp, date, item.category, item.name, item.isDone]);
  });
  
  return { status: 'success', message: 'บันทึกเช็กลิสต์สำเร็จ' };
}

function saveHealthLog(data) {
  let sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('HealthLog');
  let timestamp = new Date();
  
  // ตรวจสอบความเสี่ยง
  let riskySymptoms = ['หน้ามืด', 'อ่อนแรง', 'ใจสั่น', 'หิวมาก', 'นอนไม่หลับ'];
  if (riskySymptoms.includes(data.bodyStatus)) {
    let alertSheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('RiskAlert');
    alertSheet.appendRow([
      timestamp, data.date, 'สุขภาพ/IF', 
      'มีอาการ: ' + data.bodyStatus, 
      'วันนี้ร่างกายอาจยังไม่พร้อมสำหรับ IF เข้มงวด ควรผ่อนเป็น 12/12 ดื่มน้ำให้เพียงพอ', 
      'No'
    ]);
  }
  
  sheet.appendRow([
    timestamp, data.date, data.weight, data.waist, data.sleepHours, 
    data.sleepQuality, data.walkMinutes, data.sweetDrink, data.waterEnough, data.bodyStatus
  ]);
  
  // บันทึก IF Log คู่กัน
  let ifSheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('IFLog');
  ifSheet.appendRow([
    timestamp, data.date, data.ifType, data.eatStart, data.eatEnd, data.ifSuccess, data.bodyStatus
  ]);

  return { status: 'success', message: 'บันทึกข้อมูลสุขภาพและ IF สำเร็จ' };
}

function saveFinanceLog(data) {
  let sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('FinanceLog');
  sheet.appendRow([new Date(), data.date, data.item, data.amount, data.category, data.necessity, data.note]);
  return { status: 'success', message: 'บันทึกการเงินสำเร็จ' };
}

function saveWorkLog(data) {
  let sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('WorkLog');
  sheet.appendRow([new Date(), data.date, data.mainTask, data.completedTask, data.pendingTask, data.aiUsed, data.category]);
  return { status: 'success', message: 'บันทึกงานสำเร็จ' };
}

function saveMindLog(data) {
  let sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('MindLog');
  sheet.appendRow([new Date(), data.date, data.goodThing, data.tiredThing, data.letGoThing, data.gratitude, data.kindness, data.tomorrowFocus]);
  return { status: 'success', message: 'บันทึกความรู้สึกสำเร็จ' };
}

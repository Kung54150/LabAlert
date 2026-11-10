// ตั้งค่าการเชื่อมต่อของหน้าแอป
// apiUrl = ลิงก์ Web app ของ Apps Script (ขึ้นต้น https://script.google.com/macros/s/ ลงท้าย /exec)
// ถ้าใส่ไว้ ทุกเครื่องจะเข้าใช้งานได้เลยโดยไม่ต้องวางลิงก์ ถ้าเว้นว่าง หน้าแอปจะให้วางลิงก์ครั้งเดียวต่อเครื่อง
window.LABALERT_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/AKfycbyRh_2CtdKhgAB9Rsxh9wnutvFxw8aT9FVAY5-SLsf0sfzB3x56XdcbDDRqSEKCgkNFBw/exec',
  // Firebase Cloud Messaging (แจ้งเตือนพุชแม้ปิดหน้าแอป) ค่าชุดนี้ไม่เป็นความลับ ส่วนกุญแจ Service account เก็บใน Script Properties ของ Apps Script
  firebase: {
    apiKey: 'AIzaSyCBO5ZbMiIDsLP1UWn_l1kYKqjnehcjMCQ',
    authDomain: 'labalert-44c0f.firebaseapp.com',
    projectId: 'labalert-44c0f',
    storageBucket: 'labalert-44c0f.firebasestorage.app',
    messagingSenderId: '434650164259',
    appId: '1:434650164259:web:fb70b8aefdba63898fd596'
  },
  vapidKey: 'BA_gm_iZeDR1KleL6mbAxw1Sj53BREbA5uMmqgeZhctlWC2NpiBXpUGoMo5VCBVJH8R1ZhR97WBB7-zsNy9IhbM'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('博武教育學院推廣教育中心入口網') // 設定瀏覽器分頁標題
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // 讓手機與平板瀏覽時畫面比例正常
}
function getStaffData() {
  const sheetName = "SIT"; 

  // 直接取得目前綁定的試算表與指定的「SIT」工作表
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss ? ss.getSheetByName(sheetName) : null;
  
  if (!sheet) {
    return []; // 若找不到工作表則回傳空陣列
  }

  const data = sheet.getDataRange().getValues();
  const staffList = [];

  // 從第二列開始（i = 1），跳過第一列的標題列
  // 欄位對應：[0:照片連結, 1:姓名, 2:電子郵件, 3:聯絡電話, 4:職稱, 5:負責業務]
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1]) { // 確保「姓名」欄位有值才匯入
      staffList.push({
        photo: row[0] || '',
        name: row[1] || '',
        email: row[2] || '',
        phone: row[3] || '',
        title: row[4] || '',
        business: row[5] || ''
      });
    }
  }

  return staffList;
}
function getBannerData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("extensionbanner");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const bannerList = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
if (row[2]) {
  bannerList.push({
    name: row[0] || '',
    content: row[1] || '',
    url: row[2] || ''
  });
}
  }

  return bannerList;
}
const SPREADSHEET_ID = '1OYZtTDj9tvBtq9Fyasu4ViGQozXIpHIPQrHxvB0WYOQ'; 
// 1. 給 #list-download 抓取資料用
function getListDownloadData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('list-download');
  const rows = sheet.getDataRange().getValues();
  
  let list = [];
  for (let i = 1; i < rows.length; i++) {
    let row = rows[i];
    if (row[0]) {
      list.push({
        year: String(row[0]).trim(),
        category: String(row[1]).trim(),
        name: String(row[2]).trim(),
        url: String(row[3]).trim()
      });
    }
  }
  return list;
}

// 2. 給 #TQMF 抓取資料用
function getTQMFData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('TQMF');
  const rows = sheet.getDataRange().getValues();
  
  let list = [];
  for (let i = 1; i < rows.length; i++) {
    let row = rows[i];
    if (row[0]) {
      list.push({
        year: String(row[0]).trim(),
        category: String(row[1]).trim(),
        name: String(row[2]).trim(),
        url: String(row[3]).trim()
      });
    }
  }
  return list;
}

function getVocationalCreditData(filters) {
  try {
    const sheetName = "VocationalCredit"; // 對應試算表名稱
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn("找不到指定的工作表: " + sheetName);
      return [];
    }
    
    const rows = sheet.getDataRange().getValues();
    const records = [];
    
    // 從第 2 列開始讀取（略過標題列）
    // 欄位對應：0: 張貼日期, 1: 圖片連結, 2: 分類, 3: 標題, 4: 內容, 5: 連結
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[3]) continue; // 如果沒有標題則跳過
      
      const publishDateStr = row[0] ? new Date(row[0]) : null;
      
      records.push({
        publishDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
        rawDate: publishDateStr,
        imageUrl: row[1] || "",  // 圖片連結（供前端縮圖使用，隱藏不顯示於純文字欄位）
        category: row[2] || "",  // 分類
        title: row[3] || "",     // 標題
        content: row[4] || "",   // 內容
        link: row[5] || ""       // 連結
      });
    }
    
    // 套用前端傳來的篩選條件
    let filtered = records;
    if (filters) {
      if (filters.category && filters.category !== "全") {
        filtered = filtered.filter(item => item.category === filters.category);
      }
      
      if (filters.keyword) {
        const kw = filters.keyword.toLowerCase();
        filtered = filtered.filter(item => 
          item.title.toLowerCase().includes(kw) || item.content.toLowerCase().includes(kw)
        );
      }
      
      if (filters.startDate) {
        const start = new Date(filters.startDate);
        filtered = filtered.filter(item => item.rawDate && item.rawDate >= start);
      }
      
      if (filters.endDate) {
        const end = new Date(filters.endDate);
        end.setHours(23, 59, 59, 999);
        filtered = filtered.filter(item => item.rawDate && item.rawDate <= end);
      }
    }
    
    // 依照張貼日期由新到舊排序
    filtered.sort((a, b) => {
      if (!a.rawDate || !b.rawDate) return 0;
      return b.rawDate - a.rawDate;
    });
    
    return filtered.map(item => {
      delete item.rawDate;
      return item;
    });
    
  } catch (e) {
    console.error("getVocationalCreditData 錯誤: " + e.toString());
    return [];
  }
}

function getGovSubsidyData(filters) {
  try {
    const sheetName = "GovSubsidy"; // 對應試算表名稱
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn("找不到指定的工作表: " + sheetName);
      return [];
    }
    
    const rows = sheet.getDataRange().getValues();
    const records = [];
    
    // 從第 2 列開始讀取（略過標題列）
    // 欄位對應：0: 張貼日期, 1: 圖片連結, 2: 分類, 3: 標題, 4: 內容, 5: 連結
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[3]) continue; // 如果沒有標題則跳過
      
      const publishDateStr = row[0] ? new Date(row[0]) : null;
      
      records.push({
        publishDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
        rawDate: publishDateStr,
        imageUrl: row[1] || "",  // 圖片連結（供前端縮圖使用，隱藏不顯示於純文字欄位）
        category: row[2] || "",  // 分類
        title: row[3] || "",     // 標題
        content: row[4] || "",   // 內容
        link: row[5] || ""       // 連結
      });
    }
    
    // 套用前端傳來的篩選條件
    let filtered = records;
    if (filters) {
      if (filters.category && filters.category !== "全") {
        filtered = filtered.filter(item => item.category === filters.category);
      }
      
      if (filters.keyword) {
        const kw = filters.keyword.toLowerCase();
        filtered = filtered.filter(item => 
          item.title.toLowerCase().includes(kw) || item.content.toLowerCase().includes(kw)
        );
      }
      
      if (filters.startDate) {
        const start = new Date(filters.startDate);
        filtered = filtered.filter(item => item.rawDate && item.rawDate >= start);
      }
      
      if (filters.endDate) {
        const end = new Date(filters.endDate);
        end.setHours(23, 59, 59, 999);
        filtered = filtered.filter(item => item.rawDate && item.rawDate <= end);
      }
    }
    
    // 依照張貼日期由新到舊排序
    filtered.sort((a, b) => {
      if (!a.rawDate || !b.rawDate) return 0;
      return b.rawDate - a.rawDate;
    });
    
    return filtered.map(item => {
      delete item.rawDate;
      return item;
    });
    
  } catch (e) {
    console.error("getGovSubsidyData 錯誤: " + e.toString());
    return [];
  }
}

function getAllAnnouncementsData(filters) {
  try {
    const sheetName = "AllAnnouncements"; // 對應試算表名稱
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn("找不到指定的工作表: " + sheetName);
      return [];
    }
    
    const rows = sheet.getDataRange().getValues();
    const records = [];
    
    // 從第 2 列開始讀取（略過標題列）
    // 欄位對應：0: 張貼日期, 1: 圖片連結, 2: 分類, 3: 標題, 4: 內容, 5: 連結
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[3]) continue; // 如果沒有標題則跳過
      
      const publishDateStr = row[0] ? new Date(row[0]) : null;
      
      records.push({
        publishDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
        rawDate: publishDateStr,
        imageUrl: row[1] || "",  // 圖片連結（供前端縮圖使用，隱藏不顯示於純文字欄位）
        category: row[2] || "",  // 分類
        title: row[3] || "",     // 標題
        content: row[4] || "",   // 內容
        link: row[5] || ""       // 連結
      });
    }
    
    // 套用前端傳來的篩選條件
    let filtered = records;
    if (filters) {
      if (filters.category && filters.category !== "全") {
        filtered = filtered.filter(item => item.category === filters.category);
      }
      
      if (filters.keyword) {
        const kw = filters.keyword.toLowerCase();
        filtered = filtered.filter(item => 
          item.title.toLowerCase().includes(kw) || item.content.toLowerCase().includes(kw)
        );
      }
      
      if (filters.startDate) {
        const start = new Date(filters.startDate);
        filtered = filtered.filter(item => item.rawDate && item.rawDate >= start);
      }
      
      if (filters.endDate) {
        const end = new Date(filters.endDate);
        end.setHours(23, 59, 59, 999);
        filtered = filtered.filter(item => item.rawDate && item.rawDate <= end);
      }
    }
    
    // 依照張貼日期由新到舊排序
    filtered.sort((a, b) => {
      if (!a.rawDate || !b.rawDate) return 0;
      return b.rawDate - a.rawDate;
    });
    
    return filtered.map(item => {
      delete item.rawDate;
      return item;
    });
    
  } catch (e) {
    console.error("getAllAnnouncementsData 錯誤: " + e.toString());
    return [];
  }
}
function getScienceCreditData(filters) {
  try {
    const sheetName = "ScienceCredit"; // 對應試算表名稱
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn("找不到指定的工作表: " + sheetName);
      return [];
    }
    
    const rows = sheet.getDataRange().getValues();
    const records = [];
    
    // 從第 2 列開始讀取（略過標題列）
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[2]) continue; // 如果沒有標題則跳過
      
      const publishDateStr = row[0] ? new Date(row[0]) : null;
      
      records.push({
        publishDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
        rawDate: publishDateStr,
        category: row[1] || "",
        title: row[2] || "",
        content: row[3] || "",
        link: row[4] || "",
        fileUrl: row[5] || ""
      });
    }
    
    // 套用前端傳來的篩選條件
    let filtered = records;
    if (filters) {
      // 處理分類篩選：若有選取且不是「全」或空字串，才進行分類過濾
      if (filters.category && filters.category !== "全") {
        filtered = filtered.filter(item => item.category === filters.category);
      }
      
      if (filters.keyword) {
        const kw = filters.keyword.toLowerCase();
        filtered = filtered.filter(item => 
          item.title.toLowerCase().includes(kw) || item.content.toLowerCase().includes(kw)
        );
      }
      
      if (filters.startDate) {
        const start = new Date(filters.startDate);
        filtered = filtered.filter(item => item.rawDate && item.rawDate >= start);
      }
      
      if (filters.endDate) {
        const end = new Date(filters.endDate);
        end.setHours(23, 59, 59, 999); // 包含結束當天
        filtered = filtered.filter(item => item.rawDate && item.rawDate <= end);
      }
    }
    
    // 依照張貼日期由新到舊排序
    filtered.sort((a, b) => {
      if (!a.rawDate || !b.rawDate) return 0;
      return b.rawDate - a.rawDate;
    });
    
    // 移除暫存的 rawDate 屬性再回傳給前端
    return filtered.map(item => {
      delete item.rawDate;
      return item;
    });
    
  } catch (e) {
    console.error("getScienceCreditData 錯誤: " + e.toString());
    return [];
  }
}
function getSeniorAlbumData() {
  try {
    const sheetName = "SeniorAlbum"; // 對應試算表名稱
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn("找不到指定的工作表: " + sheetName);
      return [];
    }
    
    const rows = sheet.getDataRange().getValues();
    const records = [];
    
    // 從第 2 列開始讀取（略過標題列）
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[1]) continue; // 如果沒有標題則跳過
      
      const publishDateStr = row[0] ? new Date(row[0]) : null;
      
      records.push({
        publishDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
        rawDate: publishDateStr,
        title: row[1] || "",
        url: row[2] || ""
      });
    }
    
    // 依照張貼日期由新到舊排序
    records.sort((a, b) => {
      if (!a.rawDate || !b.rawDate) return 0;
      return b.rawDate - a.rawDate;
    });
    
    // 移除暫存的 rawDate 屬性再回傳給前端
    return records.map(item => {
      delete item.rawDate;
      return item;
    });
    
  } catch (e) {
    console.error("getSeniorAlbumData 錯誤: " + e.toString());
    return [];
  }
}
function getSeniorCoursesData(filters) {
  try {
    const sheetName = "SeniorCourses"; // 對應試算表名稱改為 SeniorCourses
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn("找不到指定的工作表: " + sheetName);
      return [];
    }
    
    const rows = sheet.getDataRange().getValues();
    const records = [];
    
    // 從第 2 列開始讀取（略過標題列）
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[2]) continue; // 如果沒有標題則跳過
      
      const publishDateStr = row[0] ? new Date(row[0]) : null;
      
      records.push({
        publishDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "",
        rawDate: publishDateStr,
        category: row[1] || "",
        title: row[2] || "",
        content: row[3] || "",
        link: row[4] || "",
        fileUrl: row[5] || ""
      });
    }
    
    // 套用前端傳來的篩選條件
    let filtered = records;
    if (filters) {
      // 處理分類篩選：若有選取且不是「全」或空字串，才進行分類過濾
      if (filters.category && filters.category !== "全") {
        filtered = filtered.filter(item => item.category === filters.category);
      }
      
      if (filters.keyword) {
        const kw = filters.keyword.toLowerCase();
        filtered = filtered.filter(item => 
          item.title.toLowerCase().includes(kw) || item.content.toLowerCase().includes(kw)
        );
      }
      
      if (filters.startDate) {
        const start = new Date(filters.startDate);
        filtered = filtered.filter(item => item.rawDate && item.rawDate >= start);
      }
      
      if (filters.endDate) {
        const end = new Date(filters.endDate);
        end.setHours(23, 59, 59, 999); // 包含結束當天
        filtered = filtered.filter(item => item.rawDate && item.rawDate <= end);
      }
    }
    
    // 依照張貼日期由新到舊排序
    filtered.sort((a, b) => {
      if (!a.rawDate || !b.rawDate) return 0;
      return b.rawDate - a.rawDate;
    });
    
    // 移除暫存的 rawDate 屬性再回傳給前端
    return filtered.map(item => {
      delete item.rawDate;
      return item;
    });
    
  } catch (e) {
    console.error("getSeniorCoursesData 錯誤: " + e.toString());
    return [];
  }
}

// ==========================================
// 🏥 प्राथमिक आरोग्य केंद्र, भादा - मुख्य बॅकएंड कोड
// ==========================================

// ⚠️ तुमचे Google Sheet आणि Drive Folders चे ID येथे टाका:
const SS_ID = "18feBIn_2h5rocFmxReqVNvI4XIeyDwOPDAxldLbadWg"; // तुमचा स्प्रेडशीट आयडी
const DOWNLOAD_FOLDER_ID = "1E6yA_nY69F6xbuw8q5xGKW7gioUuztZz"; 
const PHOTO_FOLDER_ID = "1Ga3v-sA6_AvUMzr01kIx6W1uKmBeZvRf";

const MASTER_SHEET_NAME = "MasterData";
const DATA_ENTRY_SHEET_NAME = "BsDataEntry"; 
const VILLAGE_DETAILS_SHEET_NAME = "VillageDetails"; 

function doGet() {
  return HtmlService.createTemplateFromFile('Form').evaluate()
      .setTitle('आरोग्य केंद्र डेटा सिस्टीम')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------
// 1. डेटा एन्ट्रीसाठी लागणारी माहिती (Master Data)
// ----------------------------------------------------
function getMasterData() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const values = masterSheet.getDataRange().getValues();
    const headers = values[0];
    const villageListCol = headers.indexOf("गावांची यादी"); 
    
    return values.slice(1).map(row => {
      let villages = [];
      if (villageListCol !== -1 && row[villageListCol]) {
          villages = row[villageListCol].toString().split(',').map(v => v.trim()).filter(v => v);
      }
      return {
        upkendra: row[headers.indexOf("उपकेंद्र")],
        employeeName: row[headers.indexOf("कर्मचारी नाव")],
        designation: row[headers.indexOf("पदनाम")],
        bsCode: row[headers.indexOf("Bs Code")],
        villageList: villages 
      };
    }).filter(item => item.upkendra && item.employeeName);
  } catch (e) { return []; }
}

function getLatestBundleNumber(commonDate) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const dataSheet = ss.getSheetByName(DATA_ENTRY_SHEET_NAME);
    if (!dataSheet || dataSheet.getLastRow() < 2) return 0;
    const inputDateStr = Utilities.formatDate(new Date(commonDate), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    const values = dataSheet.getDataRange().getValues();
    let latest = 0;
    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][1] instanceof Date) {
        if (Utilities.formatDate(values[i][1], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") === inputDateStr) {
          const match = values[i][6].toString().match(/(\d+)$/);
          if (match) latest = Math.max(latest, parseInt(match[1]));
        }
      }
    }
    return latest;
  } catch (e) { return 0; }
}

function getLastParayntValue(upkendra, employeeName, commonDate) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const dataSheet = ss.getSheetByName(DATA_ENTRY_SHEET_NAME);
    if (!dataSheet || dataSheet.getLastRow() < 2) return 0;
    const inputYear = new Date(commonDate).getFullYear();
    const values = dataSheet.getDataRange().getValues();
    const headers = values[0];
    const uIdx = headers.indexOf("उपकेंद्र");
    const eIdx = headers.indexOf("कर्मचारी नाव");
    const pIdx = headers.indexOf("पर्यंत (नमुना क्रमांक)");
    const dIdx = headers.indexOf("दिनांक");

    let maxParaynt = 0;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const sheetDate = new Date(row[dIdx]);
      if (isNaN(sheetDate.getTime())) continue; 
      
      if (sheetDate.getFullYear() === inputYear && row[uIdx].toString().trim() === upkendra.toString().trim() && row[eIdx].toString().trim() === employeeName.toString().trim()) {
        let currentValue = parseInt(row[pIdx]) || 0;
        if (currentValue > maxParaynt) maxParaynt = currentValue;
      }
    }
    return maxParaynt;
  } catch (e) { return 0; }
}

function processForm(formDataArray) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const dataSheet = ss.getSheetByName(DATA_ENTRY_SHEET_NAME);
    const villageSheet = ss.getSheetByName(VILLAGE_DETAILS_SHEET_NAME);
    const dataToWrite = [];
    const villageDataToWrite = [];
    let lastRow = dataSheet.getLastRow();

    formDataArray.forEach((entry) => {
      const dateObj = new Date(entry.bsSendDate);
      const dateStr = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "yyyyMMdd");
      const uniqueId = "BS_" + dateStr + "_" + (entry.bsCode || "0") + "_" + (lastRow + 1);
      const total = parseInt(entry.paraynt) - parseInt(entry.pasun) + 1;

      dataToWrite.push([uniqueId, dateObj, entry.upkendra, entry.employeeName, entry.designation, entry.bsCode, entry.bundleNumber, entry.pasun, entry.paraynt, total]);

      entry.villageDetails.forEach(v => {
        villageDataToWrite.push([uniqueId, entry.employeeName, dateObj, v.villageName, v.sampleCount, v.maleCount, v.femaleCount, entry.upkendra]);
      });
      lastRow++;
    });

    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
    villageSheet.getRange(villageSheet.getLastRow() + 1, 1, villageDataToWrite.length, villageDataToWrite[0].length).setValues(villageDataToWrite);

    return { success: true, message: "डेटा यशस्वीरित्या जतन झाला!" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ----------------------------------------------------
// 2. डॅशबोर्ड सांख्यिकी (Dashboard Stats with MonthMaster Logic)
// ----------------------------------------------------
function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const dataSheet = ss.getSheetByName("BsDataEntry");
    const monthSheet = ss.getSheetByName("MonthMaster");
    
    if (!dataSheet || !monthSheet) return { success: false, message: "Sheet not found" };

    const data = dataSheet.getDataRange().getValues();
    const monthData = monthSheet.getDataRange().getValues();
    
    const today = new Date();
    today.setHours(0,0,0,0);
    
    let startDate = new Date(today.getFullYear(), today.getMonth(), 1); 
    let endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0); 

    for (let m = 2; m < monthData.length; m++) {
      let mStart = monthData[m][1]; 
      let mEnd = monthData[m][4];   
      if (mStart instanceof Date && mEnd instanceof Date) {
        if (today >= mStart && today <= mEnd) {
          startDate = mStart;
          endDate = mEnd;
          break; 
        }
      }
    }
    
    endDate.setHours(23, 59, 59, 999);

    let stats = { total: 0, active: 0, passive: 0, activePercent: 0, passivePercent: 0 };

    for (let i = 1; i < data.length; i++) {
      let rowDate = data[i][1];
      if (rowDate instanceof Date) {
        if (rowDate >= startDate && rowDate <= endDate) {
          let count = parseInt(data[i][9]) || 0;
          let designation = String(data[i][4]).trim(); 
          stats.total += count;

          // ॲक्टिव्ह: आरोग्य सेवक, आरोग्य सेविका, आशा
          if (designation.includes("आरोग्य सेवक") || designation.includes("आरोग्य सेविका") || designation.includes("आशा")) {
            stats.active += count;
          } else {
            stats.passive += count;
          }
        }
      }
    }

    if (stats.total > 0) {
      stats.activePercent = Math.round((stats.active / stats.total) * 100);
      stats.passivePercent = Math.round((stats.passive / stats.total) * 100);
    }
    return { success: true, data: stats };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ----------------------------------------------------
// 3. डाऊनलोड्स आणि फोटो गॅलरी (Drive Integration)
// ----------------------------------------------------
function getDriveDownloads(targetFolderId) {
  try {
    const idToUse = targetFolderId || DOWNLOAD_FOLDER_ID; 
    const folder = DriveApp.getFolderById(idToUse);
    const foldersIter = folder.getFolders();
    const filesIter = folder.getFiles();
    
    const folderList = [];
    const fileList = [];
    
    while (foldersIter.hasNext()) {
      const subFolder = foldersIter.next();
      folderList.push({ id: subFolder.getId(), name: subFolder.getName() });
    }
    while (filesIter.hasNext()) {
      const file = filesIter.next();
      fileList.push({ name: file.getName(), url: file.getUrl() });
    }
    return { success: true, folders: folderList, files: fileList };
  } catch (e) { return { success: false, message: "फोल्डर सापडले नाही. कृपया Folder ID तपासा." }; }
}

function getDrivePhotos(targetFolderId) {
  try {
    const idToUse = targetFolderId || PHOTO_FOLDER_ID;
    const folder = DriveApp.getFolderById(idToUse);
    const foldersIter = folder.getFolders();
    const filesIter = folder.getFiles();
    
    const folderList = [];
    const photoList = [];
    
    while (foldersIter.hasNext()) {
      const subFolder = foldersIter.next();
      folderList.push({ id: subFolder.getId(), name: subFolder.getName() });
    }
    while (filesIter.hasNext()) {
      const file = filesIter.next();
      const mimeType = file.getMimeType();
      
      if (mimeType.indexOf('image/') !== -1) {
        const fileId = file.getId();
        const thumbUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w500";
        const highResUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=s0";
        
        const fullFileName = file.getName().split('.')[0]; 
        let nameParts = fullFileName.split('_');
        let title = nameParts[0]; 
        let desc = nameParts.length > 1 ? nameParts.slice(1).join('_') : ""; 
        
        photoList.push({ name: title, url: thumbUrl, fullUrl: highResUrl, description: desc });
      }
    }
    return { success: true, folders: folderList, files: photoList };
  } catch (e) { return { success: false, message: "फोटो फोल्डर सापडले नाही. कृपया Folder ID तपासा." }; }
}
// ----------------------------------------------------
// 4. डायनॅमिक लिंक्स (Important Links) - FIXED
// ----------------------------------------------------
function getImportantLinks() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    
    // १. शीटचे नाव स्पेससह किंवा कॅपिटल/स्मॉल असले तरी शोधून काढेल
    let sheet = null;
    const allSheets = ss.getSheets();
    for (let i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getName().trim().toLowerCase() === "importantlinks".toLowerCase()) {
        sheet = allSheets[i];
        break;
      }
    }
    
    // जर शीट तरीही सापडली नाही, तर याचा अर्थ SS_ID चुकीचा आहे.
    if (!sheet) return { success: false, message: "एरर: ImportantLinks शीट सापडली नाही. कृपया Code.gs मध्ये तुमचा SS_ID (Spreadsheet ID) बरोबर आहे का ते तपासा." };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    const links = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // जर पहिल्या कॉलममध्ये नाव असेल तरच घ्या
        links.push({
          name: String(data[i][0]).trim(),
          desc: String(data[i][1]).trim(),
          url: String(data[i][2]).trim() || "#",
          icon: String(data[i][3]).trim() || "🔗"
        });
      }
    }
    return { success: true, data: links };
  } catch (e) {
    return { success: false, message: "एरर: " + e.toString() };
  }
}

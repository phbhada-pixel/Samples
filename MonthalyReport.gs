/**
 * GOOGLE APPS SCRIPT FOR AUTOMATED HEALTH REPORTING (WEB APP VERSION)
 */

const CONFIG = {
  TEMPLATE_ID: '1hvJex623nopT-nVDTNtaGJTPmNrGuYoBHQNS7a4kb6E',
  OUTPUT_FOLDER_ID: '18XFhKzJDpWMCo5QEwE3UwVoZJE3epty3',
  SHEET_MASTER: 'MonthMaster',
  SHEET_DETAILS: 'VillageDetails'
};

// वेब ॲपमध्ये ड्रॉपडाऊन दाखवण्यासाठी महिन्यांची यादी पाठवणे
function getMonthListForWebApp() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID); // SS_ID Code.gs मधून घेतला जाईल
    const masterSheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
    const lastRow = masterSheet.getLastRow();
    if (lastRow < 3) return [];
    
    // तिसऱ्या ओळीपासून खालील सर्व महिने वाचणे
    const months = masterSheet.getRange(3, 1, lastRow - 2, 1).getDisplayValues();
    return months.map(r => r[0]).filter(m => m !== "");
  } catch (e) {
    return [];
  }
}

function parseDateSafe(value) {
  if (!value) return null;
  if (value instanceof Date) return value; 
  const str = String(value).trim();
  if (str.match(/^\d{1,2}-\d{1,2}-\d{4}$/)) {
    const parts = str.split('-');
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  if (str.includes('/')) {
    const parts = str.split('/');
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(value);
}

function cleanStr(str) {
  return String(str).replace(/\s+/g, '').toLowerCase().trim();
}

// वेब ॲपमधून महिना निवडून हे फंक्शन कॉल केले जाईल
// वेब ॲपमधून महिना निवडून हे फंक्शन कॉल केले जाईल
function generateMonthlyReportWebApp(selectedMonthDisplay) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const masterSheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
    const detailSheet = ss.getSheetByName(CONFIG.SHEET_DETAILS);
    
    const selectedStr = cleanStr(selectedMonthDisplay);
    if (!selectedStr) return { success: false, message: 'कृपया महिना निवडा.' };

    // डेटा वाचताना Values आणि DisplayValues दोन्ही घेणे
    const masterData = masterSheet.getDataRange().getValues(); 
    const masterDisplayData = masterSheet.getDataRange().getDisplayValues(); 

    const masterRows = masterData.slice(2); 
    const masterDisplayRows = masterDisplayData.slice(2); 
    
    // 'Display Value' वापरून अचूक महिना शोधणे
    let selectedIndex = -1;
    for (let i = 0; i < masterDisplayRows.length; i++) {
      if (cleanStr(masterDisplayRows[i][0]) === selectedStr) {
        selectedIndex = i;
        break;
      }
    }

    // जर महिना सापडला नाही
    if (selectedIndex === -1) {
      return { success: false, message: `महिना '${selectedMonthDisplay}' MonthMaster मध्ये सापडला नाही.` };
    }

    // सापडलेल्या इंडेक्सवरून मूळ तारखांचा रो घेणे
    const selectedRow = masterRows[selectedIndex]; 

    const fn1Start = parseDateSafe(selectedRow[1]);
    const fn1End   = parseDateSafe(selectedRow[2]);
    const fn2Start = parseDateSafe(selectedRow[3]);
    const fn2End   = parseDateSafe(selectedRow[4]);
    
    if (!fn1Start || !fn2End) return { success: false, message: "तारीख त्रुटी: MonthMaster मध्ये Start/End तारखा तपासा." };

    const reportStart = fn1Start;
    const reportEnd = fn2End;
    const ytdStart = new Date(reportEnd.getFullYear(), 0, 1);
    const ytdEnd = reportEnd;

    const detailData = detailSheet.getDataRange().getValues();
    const detailRows = detailData.slice(1);

    // Stats Object 
    const stats = {
      newOpd: selectedRow[5], newOpdYtd: 0,
      passive: 0, passiveYtd: 0,
      activeM: 0, activeF: 0, activeMYtd: 0, activeFYtd: 0,
      passiveM: 0, passiveF: 0, passiveMYtd: 0, passiveFYtd: 0,
      mpw: { fn1: 0, fn2: 0, ytd: 0, fn1Ytd: 0, fn2Ytd: 0 },
      anm: { month: 0, ytd: 0 },
      asha: { month: 0, ytd: 0 },
      colG: Number(selectedRow[6]) || 0, colGYtd: 0, 
      colH: Number(selectedRow[7]) || 0, colHYtd: 0, 
      colI: Number(selectedRow[8]) || 0, colIYtd: 0, 
      colJ: Number(selectedRow[9]) || 0, colJYtd: 0,
      d65_7days: 0, d66_15days: 0, d67_above15days: 0
    };

    masterRows.forEach(row => {
      const rDate = parseDateSafe(row[4]);
      if (rDate && rDate >= ytdStart && rDate <= ytdEnd) {
        stats.newOpdYtd += (Number(row[5]) || 0);
        stats.colGYtd += (Number(row[6]) || 0);
        stats.colHYtd += (Number(row[7]) || 0);
        stats.colIYtd += (Number(row[8]) || 0);
        stats.colJYtd += (Number(row[9]) || 0);
      }
    });

    // 1. सर्व गावांचा इतिहास आधीच गोळा करून ठेवणे
    const villageHistory = new Map();
    detailRows.forEach(row => {
      const rDate = parseDateSafe(row[2]);
      if (!rDate || isNaN(rDate.getTime())) return;
      const vKey = String(row[7]).trim() + "_" + String(row[3]).trim();
      if (!villageHistory.has(vKey)) villageHistory.set(vKey, []);
      villageHistory.get(vKey).push(rDate.getTime());
    });
    villageHistory.forEach(dates => dates.sort((a, b) => a - b));

    const villageMap = new Map();

    // 2. मुख्य कॅल्क्युलेशन लूप
    detailRows.forEach(row => {
      const rDate = parseDateSafe(row[2]);
      if (!rDate || isNaN(rDate.getTime())) return;

      const type = String(row[8]);
      const post = String(row[9]).trim();
      const male = Number(row[5]) || 0;
      const female = Number(row[6]) || 0;
      const total = Number(row[4]) || 0;
      const subCentre = String(row[7]);
      const village = String(row[3]);

      if (rDate >= ytdStart && rDate <= ytdEnd) {
        if (type === 'Passive') { stats.passiveYtd += total; stats.passiveMYtd += male; stats.passiveFYtd += female; }
        if (type === 'Active') { stats.activeMYtd += male; stats.activeFYtd += female; }
        if (post === "आरोग्य सेवक") {
          stats.mpw.ytd += total;
          if (rDate.getDate() <= 15) { stats.mpw.fn1Ytd += total; } else { stats.mpw.fn2Ytd += total; }
        }
        if (post === "आरोग्य सेविका") stats.anm.ytd += total;
        if (post === "आशा") stats.asha.ytd += total;
      }

      if (rDate >= reportStart && rDate <= reportEnd) {
        let addedToD19 = false;
        const rowTrueTotal = male + female; 

        if (type === 'Passive') { stats.passive += total; stats.passiveM += male; stats.passiveF += female; addedToD19 = true; }
        if (type === 'Active') { stats.activeM += male; stats.activeF += female; addedToD19 = true; }

        if (addedToD19) {
          const vKey = String(row[7]).trim() + "_" + String(row[3]).trim();
          const history = villageHistory.get(vKey) || [];
          let prevTime = null;
          for (let j = history.length - 1; j >= 0; j--) {
            if (history[j] < rDate.getTime()) { prevTime = history[j]; break; }
          }
          if (prevTime !== null) {
            const diffDays = Math.floor((rDate.getTime() - prevTime) / (1000 * 60 * 60 * 24));
            if (diffDays <= 7) stats.d65_7days += rowTrueTotal;
            else if (diffDays <= 15) stats.d66_15days += rowTrueTotal;
            else stats.d67_above15days += rowTrueTotal;
          } else {
            stats.d67_above15days += rowTrueTotal;
          }
        }

        if (post === "आरोग्य सेवक") {
          if (rDate >= fn1Start && rDate <= fn1End) stats.mpw.fn1 += total;
          if (rDate >= fn2Start && rDate <= fn2End) stats.mpw.fn2 += total;
        }
        if (post === "आरोग्य सेविका") stats.anm.month += total;
        if (post === "आशा") stats.asha.month += total;

        if (!villageMap.has(subCentre)) villageMap.set(subCentre, new Map());
        const vMap = villageMap.get(subCentre);
        if (!vMap.has(village)) vMap.set(village, { d50:0, d51:0, d53:0, d54:0, d56:0, d57:0, d59:0, d60:0 });
        const v = vMap.get(village);

        if (post === "आरोग्य सेवक") {
          if (rDate >= fn1Start && rDate <= fn1End) { v.d50 += male; v.d51 += female; }
          if (rDate >= fn2Start && rDate <= fn2End) { v.d53 += male; v.d54 += female; }
        } else if (post === "आरोग्य सेविका") {
          v.d56 += male; v.d57 += female;
        } else if (post === "आशा") {
          v.d59 += male; v.d60 += female;
        }
      }
    });

    const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID);
    const targetFolder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    const tempFile = templateFile.makeCopy("TEMP_REPORT", targetFolder);
    const doc = DocumentApp.openById(tempFile.getId());
    const body = doc.getBody();

    const r = {};
    const monthNames = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
    const niceMonth = monthNames[reportEnd.getMonth()] + ":" + reportEnd.getFullYear();
    r['{{D47}}'] = niceMonth;

    r['{{D1}}'] = stats.newOpd; r['{{D2}}'] = stats.newOpdYtd;
    r['{{D3}}'] = stats.passive; r['{{D4}}'] = stats.passiveYtd;
    r['{{D5}}'] = stats.activeF; r['{{D6}}'] = stats.activeM; r['{{D7}}'] = stats.activeF + stats.activeM;
    r['{{D8}}'] = stats.activeFYtd; r['{{D9}}'] = stats.activeMYtd; r['{{D10}}'] = stats.activeFYtd + stats.activeMYtd;
    r['{{D11}}'] = stats.passiveF; r['{{D12}}'] = stats.passiveM; r['{{D13}}'] = stats.passiveF + stats.passiveM;
    r['{{D14}}'] = stats.passiveFYtd; r['{{D15}}'] = stats.passiveMYtd; r['{{D16}}'] = stats.passiveFYtd + stats.passiveMYtd;
    r['{{D17}}'] = stats.activeF + stats.passiveF; r['{{D18}}'] = stats.activeM + stats.passiveM; r['{{D19}}'] = r['{{D17}}'] + r['{{D18}}'];
    r['{{D20}}'] = stats.activeFYtd + stats.passiveFYtd; r['{{D21}}'] = stats.activeMYtd + stats.passiveMYtd; r['{{D22}}'] = r['{{D20}}'] + r['{{D21}}'];
    
    r['{{D23}}'] = stats.mpw.fn1; r['{{D24}}'] = stats.mpw.fn1Ytd;
    r['{{D25}}'] = stats.mpw.fn2; r['{{D26}}'] = stats.mpw.fn2Ytd;
    r['{{D27}}'] = stats.anm.month; r['{{D28}}'] = stats.anm.ytd;
    r['{{D29}}'] = stats.asha.month; r['{{D30}}'] = stats.asha.ytd;
    
    r['{{D31}}'] = r['{{D23}}'] + r['{{D25}}'] + r['{{D27}}'] + r['{{D29}}'];
    r['{{D32}}'] = r['{{D24}}'] + r['{{D26}}'] + r['{{D28}}'] + r['{{D30}}'];
    r['{{D33}}'] = Math.round(r['{{D31}}'] * 0.6);
    r['{{D34}}'] = r['{{D31}}'] - r['{{D33}}'];

    r['{{D35}}'] = stats.colG; r['{{D36}}'] = stats.colH; r['{{D37}}'] = r['{{D35}}'] + r['{{D36}}'];
    r['{{D38}}'] = stats.colGYtd; r['{{D39}}'] = stats.colHYtd; r['{{D40}}'] = r['{{D38}}'] + r['{{D39}}'];
    r['{{D41}}'] = stats.colI; r['{{D42}}'] = stats.colJ; r['{{D43}}'] = r['{{D41}}'] + r['{{D42}}'];
    r['{{D44}}'] = stats.colIYtd; r['{{D45}}'] = stats.colJYtd; r['{{D46}}'] = r['{{D44}}'] + r['{{D45}}'];

    r['{{D65}}'] = stats.d65_7days;
    r['{{D66}}'] = stats.d66_15days;
    r['{{D67}}'] = stats.d67_above15days;

    for (const [key, val] of Object.entries(r)) {
      body.replaceText(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), String(val));
    }

    let targetTable = null;
    let patternRowIndex = -1;
    const tables = body.getTables();
    for (const t of tables) {
      for (let i = 0; i < t.getNumRows(); i++) {
        if (t.getRow(i).getText().includes('{{D48}}')) { targetTable = t; patternRowIndex = i; break; }
      }
      if (targetTable) break;
    }

    if (targetTable) {
      const patternRow = targetTable.getRow(patternRowIndex);
      const patternCopy = patternRow.copy();
      patternRow.removeFromParent();
      
      let insertIdx = patternRowIndex;
      let srNo = 1;
      const phcGrandTotal = Array(15).fill(0);

      for (const [scName, villages] of villageMap) {
        const scTotal = Array(15).fill(0);
        for (const [vName, d] of villages) {
          const vals = [
            d.d50, d.d51, (d.d50+d.d51), d.d53, d.d54, (d.d53+d.d54), 
            d.d56, d.d57, (d.d56+d.d57), d.d59, d.d60, (d.d59+d.d60), 
            (d.d50+d.d53+d.d56+d.d59), (d.d51+d.d54+d.d57+d.d60), 
            (d.d50+d.d51+d.d53+d.d54+d.d56+d.d57+d.d59+d.d60) 
          ];
          vals.forEach((v, i) => scTotal[i] += v);
          
          const newRow = targetTable.insertTableRow(insertIdx++, patternCopy.copy());
          const rowMap = {
            '{{Sr}}': srNo++, '{{D48}}': scName, '{{D49}}': vName,
            '{{D50}}': vals[0], '{{D51}}': vals[1], '{{D52}}': vals[2], '{{D53}}': vals[3], '{{D54}}': vals[4], '{{D55}}': vals[5],
            '{{D56}}': vals[6], '{{D57}}': vals[7], '{{D58}}': vals[8], '{{D59}}': vals[9], '{{D60}}': vals[10], '{{D61}}': vals[11],
            '{{D62}}': vals[12], '{{D63}}': vals[13], '{{D64}}': vals[14]
          };
          replaceInRow(newRow, rowMap);
        }
        
        const scRow = targetTable.insertTableRow(insertIdx++, patternCopy.copy());
        const scMap = {
          '{{Sr}}': '', '{{D48}}': 'Total', '{{D49}}': scName,
          '{{D50}}': scTotal[0], '{{D51}}': scTotal[1], '{{D52}}': scTotal[2], '{{D53}}': scTotal[3], '{{D54}}': scTotal[4], '{{D55}}': scTotal[5],
          '{{D56}}': scTotal[6], '{{D57}}': scTotal[7], '{{D58}}': scTotal[8], '{{D59}}': scTotal[9], '{{D60}}': scTotal[10], '{{D61}}': scTotal[11],
          '{{D62}}': scTotal[12], '{{D63}}': scTotal[13], '{{D64}}': scTotal[14]
        };
        replaceInRow(scRow, scMap);
        styleRow(scRow, true); 
        scTotal.forEach((v, i) => phcGrandTotal[i] += v);
      }
      
      const gtRow = targetTable.insertTableRow(insertIdx++, patternCopy.copy());
      const gtMap = {
        '{{Sr}}': '', '{{D48}}': 'PHC', '{{D49}}': 'Grand Total',
        '{{D50}}': phcGrandTotal[0], '{{D51}}': phcGrandTotal[1], '{{D52}}': phcGrandTotal[2], '{{D53}}': phcGrandTotal[3], '{{D54}}': phcGrandTotal[4], '{{D55}}': phcGrandTotal[5],
        '{{D56}}': phcGrandTotal[6], '{{D57}}': phcGrandTotal[7], '{{D58}}': phcGrandTotal[8], '{{D59}}': phcGrandTotal[9], '{{D60}}': phcGrandTotal[10], '{{D61}}': phcGrandTotal[11],
        '{{D62}}': phcGrandTotal[12], '{{D63}}': phcGrandTotal[13], '{{D64}}': phcGrandTotal[14]
      };
      replaceInRow(gtRow, gtMap);
      styleRow(gtRow, true);
    }

    doc.saveAndClose();
    Utilities.sleep(3000); 
    
    const pdfBlob = tempFile.getAs(MimeType.PDF).setName(`Report_${niceMonth.replace(':', '_')}.pdf`);
    const fileUrl = targetFolder.createFile(pdfBlob).getUrl(); 
    tempFile.setTrashed(true);
    
    return { success: true, message: "PDF रिपोर्ट यशस्वीरित्या तयार झाला!", url: fileUrl };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
function replaceInRow(row, map) {
  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    let text = cell.getText();
    for (const [key, val] of Object.entries(map)) {
      if (text.includes(key)) {
        cell.replaceText(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), String(val));
      }
    }
  }
}

function styleRow(row, bold) {
  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    cell.editAsText().setBold(bold); 
    if(bold) cell.setBackgroundColor('#F3F3F3');
  }
}

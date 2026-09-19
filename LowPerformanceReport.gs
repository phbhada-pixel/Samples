/**
 * प्राथमिक आरोग्य केंद्र भादा - कमी कामगिरी अहवाल (निरंक व 75% पेक्षा कमी)
 */
function generateLowPerformanceReportWebApp(selectedMonthDisplay) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const monthMasterSheet = ss.getSheetByName("MonthMaster");
    const empMasterSheet = ss.getSheetByName("MasterData");
    const bsEntrySheet = ss.getSheetByName("BsDataEntry");
    
    const MONTHLY_TARGET = 50; 
    const TARGET_75_PERCENT = MONTHLY_TARGET * 0.75; // 37.5 नमुने

    if (!monthMasterSheet || !empMasterSheet || !bsEntrySheet) {
      return { success: false, message: "Sheet सापडली नाही (MonthMaster, MasterData किंवा BsDataEntry)." };
    }

    const selectedStr = String(selectedMonthDisplay).replace(/\s+/g, '').toLowerCase().trim();
    if (!selectedStr) return { success: false, message: 'कृपया अहवालासाठी महिना निवडा.' };

    // १. MonthMaster मधून निवडलेल्या महिन्याच्या तारखा काढणे
    const masterDisplayData = monthMasterSheet.getDataRange().getDisplayValues(); 
    const masterRows = monthMasterSheet.getDataRange().getValues().slice(2); 
    const masterDisplayRows = masterDisplayData.slice(2); 
    
    let selectedIndex = -1;
    for (let i = 0; i < masterDisplayRows.length; i++) {
      if (String(masterDisplayRows[i][0]).replace(/\s+/g, '').toLowerCase().trim() === selectedStr) {
        selectedIndex = i; break;
      }
    }

    if (selectedIndex === -1) return { success: false, message: `महिना '${selectedMonthDisplay}' MonthMaster मध्ये सापडला नाही.` };

    const selectedMonthRow = masterRows[selectedIndex]; 
    
    function parseDateSafe(value) {
      if (!value) return null;
      if (value instanceof Date) return value; 
      const str = String(value).trim();
      if (str.match(/^\d{1,2}-\d{1,2}-\d{4}$/)) { const parts = str.split('-'); return new Date(parts[2], parts[1] - 1, parts[0]); }
      if (str.includes('/')) { const parts = str.split('/'); return new Date(parts[2], parts[1] - 1, parts[0]); }
      return new Date(value);
    }

    const f1Start = parseDateSafe(selectedMonthRow[1]); 
    let f2End = parseDateSafe(selectedMonthRow[4]); 
    if (!f1Start || !f2End) return { success: false, message: "तारीख त्रुटी: MonthMaster मध्ये तारखा तपासा." };
    f2End.setHours(23, 59, 59, 999);

    const empData = empMasterSheet.getDataRange().getValues();
    const bsData = bsEntrySheet.getDataRange().getValues(); 

    let zeroSamplesList = [];
    let lowSamplesList = [];

    // २. कर्मचाऱ्यांची कामगिरी तपासणे
    for (let k = 1; k < empData.length; k++) {
      let scName = empData[k][0]; 
      let empName = empData[k][1]; 
      let post = empData[k][2] ? empData[k][2].toString().trim() : ""; 
      let bsCode = empData[k][3];  
      
      // 'आशा' कर्मचारी वगळा
      if (!empName || !bsCode || post.includes("आशा")) continue;

      let totalSamplesSum = 0; 
      for (let j = 1; j < bsData.length; j++) {
        let rowDate = new Date(bsData[j][1]); 
        let rowBsCode = bsData[j][5];         
        let rowSamples = Number(bsData[j][9]) || 0; 

        if (rowBsCode == bsCode && rowDate >= f1Start && rowDate <= f2End) {
          totalSamplesSum += rowSamples; 
        }
      }

      // कॅटेगरीनुसार विभागणी
      if (totalSamplesSum === 0) {
        zeroSamplesList.push({ name: empName, post: post, sc: scName, count: totalSamplesSum });
      } else if (totalSamplesSum < TARGET_75_PERCENT) {
        lowSamplesList.push({ name: empName, post: post, sc: scName, count: totalSamplesSum });
      }
    }

    // ३. PDF साठी HTML तयार करणे
    if (zeroSamplesList.length === 0 && lowSamplesList.length === 0) {
      return { success: true, message: "उत्तम! निवडलेल्या महिन्यात सर्व कर्मचाऱ्यांनी ७५% पेक्षा जास्त काम केले आहे. कोणीही निरंक नाही." };
    }

    let reportDate = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy");
    let mainFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    
    let html = `<html><head><style>
      @page { size: A4; margin: 1.5cm; }
      body { font-family: 'Arial'; color: #000; margin: 0; padding: 0; font-size: 14px; }
      .header { text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
      .title { font-size: 18px; font-weight: bold; margin: 5px 0; }
      .sub-title { font-size: 16px; margin: 5px 0; text-decoration: underline; }
      .section-title { font-size: 15px; font-weight: bold; color: #b71c1c; margin-top: 25px; margin-bottom: 10px; }
      .table-box { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
      .table-box th { background-color: #f2f2f2; border: 1px solid #000; padding: 8px; font-weight: bold; text-align: center; }
      .table-box td { border: 1px solid #000; padding: 6px; text-align: center; }
      .text-left { text-align: left !important; }
      .footer-sign { margin-top: 50px; text-align: right; font-weight: bold; }
    </style></head><body>
      
      <div class="header">
        <div class="title">प्राथमिक आरोग्य केंद्र भादा, ता. औसा जि. लातूर</div>
        <div class="sub-title">कमी कामगिरी अहवाल (रक्त नमुना संकलन) - माहे: ${selectedMonthDisplay}</div>
        <div style="text-align: right; font-size: 12px; margin-top: 5px;">अहवाल दिनांक: ${reportDate}</div>
      </div>

      <p>माहे ${selectedMonthDisplay} मध्ये <b>५०</b> नमुन्यांच्या उद्दिष्टापैकी निरंक (०) आणि ७५% पेक्षा कमी (३७ पेक्षा कमी) नमुने घेतलेल्या कर्मचाऱ्यांची यादी खालीलप्रमाणे आहे:</p>`;

    // निरंक यादी टेबल
    if (zeroSamplesList.length > 0) {
      html += `<div class="section-title">१. निरंक (०) नमुने घेतलेले कर्मचारी:</div>
      <table class="table-box">
        <tr><th style="width:10%;">अ.क्र.</th><th style="width:40%;">कर्मचाऱ्याचे नाव (पद)</th><th style="width:30%;">उपकेंद्र</th><th style="width:20%;">घेतलेले नमुने</th></tr>`;
      zeroSamplesList.forEach((emp, i) => {
        html += `<tr><td>${i+1}</td><td class="text-left">${emp.name} (${emp.post})</td><td>${emp.sc}</td><td style="color:red; font-weight:bold;">${emp.count}</td></tr>`;
      });
      html += `</table>`;
    }

    // 75% पेक्षा कमी यादी टेबल
    if (lowSamplesList.length > 0) {
      html += `<div class="section-title">२. ७५% पेक्षा कमी (१ ते ३७) नमुने घेतलेले कर्मचारी:</div>
      <table class="table-box">
        <tr><th style="width:10%;">अ.क्र.</th><th style="width:40%;">कर्मचाऱ्याचे नाव (पद)</th><th style="width:30%;">उपकेंद्र</th><th style="width:20%;">घेतलेले नमुने</th></tr>`;
      lowSamplesList.forEach((emp, i) => {
        html += `<tr><td>${i+1}</td><td class="text-left">${emp.name} (${emp.post})</td><td>${emp.sc}</td><td style="font-weight:bold;">${emp.count}</td></tr>`;
      });
      html += `</table>`;
    }

    html += `<div class="footer-sign">
        वैद्यकीय अधिकारी<br>प्राथमिक आरोग्य केंद्र भादा
      </div>
    </body></html>`;

    let finalBlob = Utilities.newBlob(html, "text/html", "Low_Performance.html");
    let pdfFile = mainFolder.createFile(finalBlob.getAs("application/pdf"))
                            .setName("कमी_कामगिरी_अहवाल_" + selectedMonthDisplay + ".pdf");

    return { success: true, message: `अहवाल तयार! निरंक: ${zeroSamplesList.length} आणि कमी काम: ${lowSamplesList.length} कर्मचारी सापडले.`, url: pdfFile.getUrl() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

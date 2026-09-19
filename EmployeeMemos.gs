/**
 * प्राथमिक आरोग्य केंद्र भादा - कर्मचारी कामगिरी आणि एकत्रित नोटीस प्रणाली (चेकबॉक्स निवडीसह)
 */

// --- १. फक्त पात्र कर्मचाऱ्यांची यादी शोधून पाठवणारे फंक्शन ---
function getDefaulterListForMonth(selectedMonthDisplay) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const monthMasterSheet = ss.getSheetByName("MonthMaster");
    const empMasterSheet = ss.getSheetByName("MasterData");
    const bsEntrySheet = ss.getSheetByName("BsDataEntry");
    
    const MONTHLY_TARGET = 50; 
    if (!monthMasterSheet || !empMasterSheet || !bsEntrySheet) return { success: false, message: "Sheets सापडल्या नाहीत." };

    const selectedStr = String(selectedMonthDisplay).replace(/\s+/g, '').toLowerCase().trim();
    if (!selectedStr) return { success: false, message: 'महिना रिक्त आहे.' };

    const masterDisplayRows = monthMasterSheet.getDataRange().getDisplayValues().slice(2); 
    const masterRows = monthMasterSheet.getDataRange().getValues().slice(2); 
    
    let selectedIndex = -1;
    for (let i = 0; i < masterDisplayRows.length; i++) {
      if (String(masterDisplayRows[i][0]).replace(/\s+/g, '').toLowerCase().trim() === selectedStr) { selectedIndex = i; break; }
    }
    if (selectedIndex === -1) return { success: false, message: `महिना सापडला नाही.` };

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
    if (!f1Start || !f2End) return { success: false, message: "MonthMaster मध्ये तारखांची त्रुटी." };
    f2End.setHours(23, 59, 59, 999);

    const empData = empMasterSheet.getDataRange().getValues();
    const bsData = bsEntrySheet.getDataRange().getValues(); 

    let combinedList = [];

    for (let k = 1; k < empData.length; k++) {
      let scName = empData[k][0]; 
      let empName = empData[k][1]; 
      let post = empData[k][2] ? empData[k][2].toString().trim() : ""; 
      let bsCode = empData[k][3];  
      if (!empName || !bsCode || post.includes("आशा")) continue;

      let totalSamplesSum = 0; 
      let uniqueDays = new Set(); 
      for (let j = 1; j < bsData.length; j++) {
        let rowDate = new Date(bsData[j][1]); 
        let rowBsCode = bsData[j][5];         
        let rowSamples = Number(bsData[j][9]) || 0; 
        if (rowBsCode == bsCode && rowDate >= f1Start && rowDate <= f2End) {
          totalSamplesSum += rowSamples; uniqueDays.add(Utilities.formatDate(rowDate, "GMT+5:30", "yyyy-MM-dd"));
        }
      }
      if (totalSamplesSum < MONTHLY_TARGET) {
        combinedList.push({ name: empName, post: post, sc: scName, daysCount: uniqueDays.size, samplesCount: totalSamplesSum });
      }
    }
    return { success: true, defaulters: combinedList };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// --- २. निवडलेल्या कर्मचाऱ्यांसाठी PDF नोटीस तयार करणारे फंक्शन ---
function generateEmployeeNoticesWebApp(selectedMonthDisplay, selectedEmpNames) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const MONTHLY_TARGET = 50; 
    
    // १. पुन्हा तीच यादी काढणे (डेटाची सुरक्षितता)
    const listResponse = getDefaulterListForMonth(selectedMonthDisplay);
    if(!listResponse.success) return listResponse;
    
    let combinedList = listResponse.defaulters;

    // २. फ्रंटएंडवरून आलेल्या नावांनुसार फिल्टर करणे
    combinedList = combinedList.filter(e => selectedEmpNames.indexOf(e.name) !== -1);

    if (combinedList.length === 0) {
      return { success: false, message: "नोटीस तयार करण्यासाठी कोणतेही कर्मचारी निवडले नाहीत." };
    }

    let noticeDate = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy");
    let mainFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    
    let allNoticesHtml = `<html><head><style>
      @page { size: A4; margin: 1cm; }
      body { font-family: 'Arial'; line-height: 1.4; color: #000; margin: 0; padding: 0; }
      .notice-container { position: relative; min-height: 26cm; border-bottom: 1px solid #ccc; padding-bottom: 10px; }
      .header { text-align: center; border-bottom: 2px solid #000; padding-bottom: 5px; margin-bottom: 10px; }
      .table-box { width: 100%; border-collapse: collapse; margin: 15px 0; }
      .table-box td { border: 1px solid #000; padding: 8px; font-size: 14px; }
      .page-break { page-break-after: always; }
      .footer-sign { margin-top: 30px; text-align: right; }
      .copy-to { margin-top: 40px; font-size: 12px; border-top: 1px dashed #000; padding-top: 10px; }
    </style></head><body>`;

    combinedList.forEach((e, index) => {
      allNoticesHtml += `
      <div class="notice-container">
        <div class="header">
          <h2 style="margin:0; font-size: 20px;">प्राथमिक आरोग्य केंद्र भादा, ता. औसा जि. लातूर</h2>
          <p style="margin:5px 0;"><b>कारणे दाखवा नोटीस</b></p>
        </div>
        
        <p style="margin: 0;"><b>जा.क्र. प्रआकेभा/कार्यालय/२०२६/........</b> <span style="float:right;"><b>दिनांक: ${noticeDate}</b></span></p>
        <br>
        <p style="margin: 5px 0;"><b>प्रति,</b><br>श्रीमती/श्री. ${e.name}<br>पद: ${e.post}, उपकेंद्र: ${e.sc}</p>
        
        <p style="margin: 15px 0;"><b>विषय: माहे ${selectedMonthDisplay} मधील रक्त नमुना संकलनाचे उद्दिष्ट पूर्ण न केल्याबाबत...</b></p>
        
        <p>महोदय/महोदया,</p>
        <p>आपणास कळविण्यात येते की, राष्ट्रीय आरोग्य कार्यक्रमांतर्गत हिवताप दुरीकरणासाठी दरमहा <b>५०</b> रक्त नमुने संकलित करण्याचे उद्दिष्ट आपणास देण्यात आलेले आहे. माहे ${selectedMonthDisplay} मधील आपल्या कामगिरीचा अहवाल खालीलप्रमाणे आहे:</p>
        
        <table class="table-box">
          <tr><td style="width:60%;">निर्धारित मासिक उद्दिष्ट</td><td><b>${MONTHLY_TARGET} नमुने</b></td></tr>
          <tr><td>आपण प्रत्यक्षात घेतलेले एकूण नमुने</td><td><b style="color:red;">${e.samplesCount} नमुने</b></td></tr>
          <tr><td>नमुना संकलन केलेले एकूण दिवस</td><td><b>${e.daysCount} दिवस</b></td></tr>
        </table>

        <p>वरील विवरणावरून असे निदर्शनास येते की, आपल्याकडून निर्धारित उद्दिष्टाच्या तुलनेत अत्यंत कमी कामगिरी झाली आहे. तरी, याबाबतचा आपला लेखी खुलासा २ दिवसांच्या आत प्रत्यक्ष सादर करावा.</p>
        
        <div class="footer-sign">
          <br><br>
          <b>वैद्यकीय अधिकारी</b><br>प्राथमिक आरोग्य केंद्र भादा
        </div>

        <div class="copy-to">
          <b>प्रत माहितीस्तव सविनय सादर:</b><br>
          मा. तालुका आरोग्य अधिकारी, तालुका आरोग्य कार्यालय औसा, जि. लातूर.<br>
          <div style="text-align: right; margin-top: 20px;">
            <b>वैद्यकीय अधिकारी</b><br>प्राथमिक आरोग्य केंद्र भादा
          </div>
        </div>
      </div>`;
      
      if (index < combinedList.length - 1) {
        allNoticesHtml += '<div class="page-break"></div>';
      }
    });

    allNoticesHtml += "</body></html>";

    let finalBlob = Utilities.newBlob(allNoticesHtml, "text/html", "Notices.html");
    let pdfFile = mainFolder.createFile(finalBlob.getAs("application/pdf"))
                            .setName("एकत्रित_नोटीस_भादा_" + selectedMonthDisplay + ".pdf");

    return { success: true, message: `यशस्वी! निवडलेल्या ${combinedList.length} कर्मचाऱ्यांची PDF नोटीस तयार झाली आहे.`, url: pdfFile.getUrl() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

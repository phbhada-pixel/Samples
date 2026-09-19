/**
 * चालू वर्षाची जुनी फाईल शोधून हटवणे आणि नवीन PDF तयार करणे (फिल्टरसह)
 */
function generateEmployeeRegisterWebApp(filterUpkendra, filterEmployee) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sourceSheet = ss.getSheetByName("BsDataEntry");
    if (!sourceSheet) return { success: false, message: "एरर: BsDataEntry शीट सापडली नाही." };

    var values = sourceSheet.getDataRange().getValues();
    var dataRows = values.slice(1);
    var now = new Date();
    var currentYear = now.getFullYear();
    
    // १. चालू वर्षाचा आणि युजरने निवडलेल्या फिल्टरनुसार डेटा शोधणे
    var filteredRows = dataRows.filter(function(row) {
      var rowDate = new Date(row[1]);
      var isCurrentYear = !isNaN(rowDate.getTime()) && rowDate.getFullYear() === currentYear;
      if (!isCurrentYear) return false;

      // उपकेंद्र फिल्टर (जर 'All' नसेल तरच तपासा)
      if (filterUpkendra && filterUpkendra !== "All" && row[2] !== filterUpkendra) return false;

      // कर्मचारी फिल्टर (जर 'All' नसेल तरच तपासा)
      if (filterEmployee && filterEmployee !== "All" && row[3] !== filterEmployee) return false;

      return true;
    });

    if (filteredRows.length === 0) {
      return { success: false, message: "निवडलेल्या उपकेंद्र/कर्मचाऱ्यासाठी चालू वर्षाचा कोणताही डेटा सापडला नाही." };
    }

    // २. डेटा गटबद्ध करणे (Grouping)
    var organizedData = {};
    filteredRows.forEach(function(row) {
      var subcenter = row[2];
      var employee = row[3];
      if (!organizedData[subcenter]) organizedData[subcenter] = {};
      if (!organizedData[subcenter][employee]) {
        organizedData[subcenter][employee] = { designation: row[4], bsCode: row[5], entries: [] };
      }
      organizedData[subcenter][employee].entries.push(row);
    });

    // ३. फोल्डर व्यवस्थापन
    var fileId = ss.getId();
    var parentFolder = DriveApp.getFileById(fileId).getParents().next();
    var monthName = Utilities.formatDate(now, "GMT+5:30", "MMMM yyyy");
    var monthlyFolder = parentFolder.getFoldersByName(monthName).hasNext() ? 
                        parentFolder.getFoldersByName(monthName).next() : parentFolder.createFolder(monthName);

    // फाईलचे नाव ठरवणे (फिल्टरनुसार)
    var namePrefix = "BS_Register_" + currentYear;
    if (filterUpkendra !== "All") namePrefix += "_" + filterUpkendra;
    if (filterEmployee !== "All") namePrefix += "_" + filterEmployee;
    var finalPdfName = namePrefix + "_" + Utilities.formatDate(now, "GMT+5:30", "dd-MM-yyyy") + ".pdf";

    // जर त्याच नावाची जुनी फाईल असेल तर ती हटवणे
    var files = monthlyFolder.getFilesByName(finalPdfName);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }

    // ४. तात्पुरती वेगळी फाईल तयार करणे
    var tempSS = SpreadsheetApp.create("Report_Temp_" + now.getTime());
    var tempSSId = tempSS.getId();
    var subcenterKeys = Object.keys(organizedData).sort();

    subcenterKeys.forEach(function(subcenterName, index) {
      var reportSheet = (index === 0) ? tempSS.getSheets()[0] : tempSS.insertSheet();
      reportSheet.setName(subcenterName.substring(0, 30));
      
      reportSheet.getRange(1, 1, 1, 6).merge().setValue("प्राथमिक आरोग्य केंद्र भादा")
                 .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
      reportSheet.getRange(2, 1, 1, 6).merge().setValue("कर्मचारी निहाय नोंदवही (" + subcenterName + " - " + currentYear + ")")
                 .setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center");
      
      var currentRow = 5;
      var employeesInSubcenter = organizedData[subcenterName];
      
      Object.keys(employeesInSubcenter).sort().forEach(function(empName) {
        var group = employeesInSubcenter[empName];
        
        var infoText = "कर्मचारी: " + empName + "  |  पद: " + group.designation + "  |  Code: " + group.bsCode;
        reportSheet.getRange(currentRow, 1, 1, 6).merge().setValue(infoText)
                   .setFontWeight("bold").setBackground("#f3f3f3").setBorder(true, true, true, true, true, true);

        var headers = [["अ.क्र", "दिनांक", "पासून", "पर्यंत", "एकूण", "बंडल क्र."]];
        reportSheet.getRange(currentRow + 1, 1, 1, 6).setValues(headers)
                   .setBackground("#eeeeee").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

        var tableData = group.entries.sort((a,b) => new Date(a[1]) - new Date(b[1])).map(function(r, idx) {
          var dateVal = r[1] instanceof Date ? Utilities.formatDate(r[1], "GMT+5:30", "dd/MM/yyyy") : r[1];
          return [idx + 1, dateVal, r[7], r[8], r[9], r[6]];
        });

        reportSheet.getRange(currentRow + 2, 1, tableData.length, 6).setValues(tableData)
                   .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
        
        currentRow += tableData.length + 5;
      });

      reportSheet.setColumnWidth(1, 60); reportSheet.setColumnWidth(2, 90);
      reportSheet.setColumnWidth(3, 125); reportSheet.setColumnWidth(4, 125);
      reportSheet.setColumnWidth(5, 125); reportSheet.setColumnWidth(6, 100);
    });

    SpreadsheetApp.flush();

    // ५. PDF एक्सपोर्ट
    var url = "https://docs.google.com/spreadsheets/d/" + tempSSId + "/export?exportFormat=pdf&format=pdf" +
              "&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false";
    
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    
    var savedPdfFile = monthlyFolder.createFile(response.getBlob().setName(finalPdfName));
    var fileUrl = savedPdfFile.getUrl();

    // ६. स्वच्छता
    DriveApp.getFileById(tempSSId).setTrashed(true);

    return { success: true, message: "नोंदवही PDF यशस्वीरित्या तयार झाली!", url: fileUrl };

  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

/**
 * वेब ॲपसाठी: 'BsDataEntry' मधून डेटा वाचून दैनिक PDF रिपोर्ट तयार करते.
 */
function generateDailyMalariaReportWebApp(dateStr) {
  
  // --- कॉन्फिगरेशन ---
  // (SS_ID तुमच्या Code.gs मध्ये आधीच आहे, त्यामुळे तो ग्लोबल वापरला जाईल)
  const DATA_SHEET_NAME = 'BsDataEntry'; 
  
  // कॉलम इंडेक्स (0-आधारित)
  const COL_BS_SEND_DATE = 1;      // दिनांक
  const COL_UPAKENDRA = 2;         // उपकेंद्र
  const COL_KARMACHARI = 3;        // कर्मचारी नाव
  const COL_DESIGNATION = 4;       // पदनाम
  const COL_BS_CODE = 5;           // Bs Code
  const COL_BUNDLE_NO = 6;         // बंडल क्रमांक
  const COL_FROM = 7;              // पासून (नमुना क्रमांक)
  const COL_TO = 8;                // पर्यंत (नमुना क्रमांक)
  const COL_TOTAL = 9;             // एकूण नमुने
  // --- कॉन्फिगरेशन समाप्त ---
  
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const timeZone = ss.getSpreadsheetTimeZone();
    
    if (!dataSheet) return { success: false, message: `एरर: ${DATA_SHEET_NAME} शीट सापडली नाही.` };
    
    // युजरने पाठवलेली तारीख Date Object मध्ये बदलणे
    const parts = dateStr.split('-');
    const reportDateValue = new Date(parts[0], parts[1] - 1, parts[2]);
    const reportDateString = Utilities.formatDate(reportDateValue, timeZone, 'yyyy-MM-dd');
    
    // डेटा फिल्टर करा
    const values = dataSheet.getDataRange().getValues();
    if (values.length <= 1) return { success: false, message: "शीटमध्ये कोणताही डेटा नाही." };
    
    const dataRows = values.slice(1); 
    let filteredData = dataRows.filter(row => {
      const bsDate = row[COL_BS_SEND_DATE];
      if (bsDate instanceof Date) {
        return Utilities.formatDate(bsDate, timeZone, 'yyyy-MM-dd') === reportDateString;
      }
      return false;
    });
    
    if (filteredData.length === 0) {
      return { success: false, message: `दिनांक ${formattedDateDisplay(reportDateValue, timeZone)} साठी कोणताही डेटा सापडला नाही.` };
    }
    
    // डेटा सॉर्ट करा (उपकेंद्र > कर्मचारी > पासून)
    filteredData.sort((a, b) => {
      if (a[COL_UPAKENDRA] !== b[COL_UPAKENDRA]) return a[COL_UPAKENDRA].localeCompare(b[COL_UPAKENDRA]);
      if (a[COL_KARMACHARI] !== b[COL_KARMACHARI]) return a[COL_KARMACHARI].localeCompare(b[COL_KARMACHARI]);
      return (parseInt(a[COL_FROM]) || 0) - (parseInt(b[COL_FROM]) || 0);
    });
    
    // HTML अहवाल तयार करा
    const colMap = {
      upakendra: COL_UPAKENDRA, name: COL_KARMACHARI, desig: COL_DESIGNATION, code: COL_BS_CODE,
      bundle: COL_BUNDLE_NO, from: COL_FROM, to: COL_TO, total: COL_TOTAL
    };

    const reportHtml = generateHtmlReport(reportDateString, filteredData, timeZone, colMap); 
    
    // PDF सेव्ह करा आणि URL मिळवा
    const folder = createFolderStructure(ss.getId(), reportDateValue, timeZone);
    const fileName = `Malaria_Report_${reportDateString}.pdf`;
    const blob = Utilities.newBlob(reportHtml, MimeType.HTML, 'report.html').getAs('application/pdf');
    const file = folder.createFile(blob.setName(fileName));
    
    return { success: true, message: "दैनिक पत्र यशस्वीरित्या तयार झाले!", url: file.getUrl() };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function formattedDateDisplay(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'dd-MM-yyyy');
}

function generateHtmlReport(dateStr, filteredData, timeZone, colMap) {
  const formattedDate = Utilities.formatDate(new Date(dateStr), timeZone, 'dd-MM-yyyy');
  let grandTotal = 0;
  let html = `<html><head><style>
    @page { size: A4; margin: 10mm; }
    body { font-family: 'Times New Roman', serif; font-size: 11pt; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th, td { border: 1px solid black; padding: 6px; text-align: center; font-size: 9pt; }
    th { background-color: #f2f2f2; }
    .text-left { text-align: left; }
    .header-info { text-align: right; font-weight: bold; }
    .subject { font-weight: bold; text-decoration: underline; margin: 15px 0; }
    .sep { background-color: #333; height: 2px; padding: 0 !important; }
  </style></head><body>
    <div class="header-info">दिनांक: ${formattedDate}</div>
    <p>प्रति,<br>प्रयोगशाळा वैज्ञानिक अधिकारी,<br>जिल्हा हिवताप अधिकारी कार्यालय लातूर</p>
    <div class="subject">विषय:- हिवताप रक्त नमुने तपासणीसाठी पाठवीत असले बाबत.</div>
    <p>महोदय, उपरोक्त विषयान्वये प्राथमिक आरोग्य केंद्र भादा अंतर्गत गोळा केलेले एकूण <b>${filteredData.length}</b> नोंदींचे नमुने सादर करीत आहोत.</p>
    <table>
      <thead>
        <tr>
          <th>अ.क्र.</th>
          <th>बंडल क्र.</th>
          <th>कर्मचारी नाव (पदनाम)</th>
          <th>उपकेंद्र</th>
          <th>BS Code</th>
          <th>पासून</th>
          <th>पर्यंत</th>
          <th>एकूण</th>
        </tr>
      </thead>
      <tbody>`;

  let lastUpakendra = null;

  filteredData.forEach((row, i) => {
    const total = parseFloat(row[colMap.total]) || 0;
    grandTotal += total;

    if (lastUpakendra !== null && row[colMap.upakendra] !== lastUpakendra) {
      html += `<tr class="sep"><td colspan="8"></td></tr>`;
    }
    lastUpakendra = row[colMap.upakendra];

    html += `<tr>
      <td>${i + 1}</td>
      <td>${row[colMap.bundle]}</td>
      <td class="text-left">${row[colMap.name]} (${row[colMap.desig]})</td>
      <td>${row[colMap.upakendra]}</td>
      <td>${row[colMap.code]}</td>
      <td>${row[colMap.from]}</td>
      <td>${row[colMap.to]}</td>
      <td>${total}</td>
    </tr>`;
  });

  html += `<tr style="background:#eee; font-weight:bold;">
    <td colspan="7" style="text-align:right">एकूण (Grand Total):</td>
    <td>${grandTotal}</td>
  </tr></tbody></table>
  <div style="margin-top:40px; text-align:right;">
    आपला विश्वासू,<br><br><br>वैद्यकीय अधिकारी,<br>प्रा. आ. केंद्र भादा
  </div>
  </body></html>`;
  
  return html;
}

function createFolderStructure(spreadsheetId, date, timeZone) {
  const file = DriveApp.getFileById(spreadsheetId);
  const parent = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  const year = Utilities.formatDate(date, timeZone, 'yyyy');
  const month = Utilities.formatDate(date, timeZone, 'MMMM');
  
  const getFolder = (p, name) => p.getFoldersByName(name).hasNext() ? p.getFoldersByName(name).next() : p.createFolder(name);
  
  const yFolder = getFolder(parent, year);
  return getFolder(yFolder, month);
}

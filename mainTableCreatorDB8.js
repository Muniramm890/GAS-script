function setupDB8MainData () {
  var db8 = getDB8(); 
  
  // 1. DB0 se School ki saari detail fetch karo
  var config = getMasterConfig(); 
  if (!config) {
    Logger.log("❌ DB0 se configuration nahi mili. Pehle DB0 setup karein.");
    return;
  }

  // 2. Columns define karo (Data Row 6 se start hoga)
  var profileStructures = [
    {
      name: "Student_Data",
      cols: ["UID", "Roll", "Name", "DOB", "Gender", "Class", "Section", "AdmDate", "Father", "Mother", "Mobile", "Email", "Address", "Password", "Photo"]
    },
    {
      name: "Teacher_Data",
      cols: ["UID", "Name", "DOB", "Gender", "Father_Spouse", "Mother", "Designation", "Subject", "Qualification", "JoinDate", "Salary", "Mobile", "Email", "Address", "Password", "Photo"]
    }
  ];

  // 3. Har profile ke liye sheet generate aur design karo
  profileStructures.forEach(function(prof) {
    var sheet = db8.getSheetByName(prof.name);
    
    // Agar sheet nahi hai toh banao, hai toh clear karke fresh start karo
    if (!sheet) {
      sheet = db8.insertSheet(prof.name);
    } else {
      sheet.clear(); 
    }

    var maxCol = prof.cols.length; // Student ke 15, Teacher ke 16 columns
    
    // ==========================================
    // 🎨 DYNAMIC HEADER GENERATION (ROW 1 TO 5)
    // ==========================================
    
    // Row 1: SCHOOL NAME (DB0 se)
    sheet.getRange(1, 1, 1, maxCol).merge().setValue(config.schoolNameBig.toUpperCase())
         .setBackground("#0c343d").setFontColor("#ffffff").setFontSize(16)
         .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Row 2: ADDRESS & AFFILIATION (DB0 se)
    var subHeader = config.address + " | Affiliation No: " + config.affiliation;
    sheet.getRange(2, 1, 1, maxCol).merge().setValue(subHeader)
         .setBackground("#e0f2f1").setFontColor("#000000").setFontSize(10)
         .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Row 3: SHEET TITLE (Student Data / Teacher Data)
    var titleText = prof.name.replace("_", " ").toUpperCase() + " REGISTRY";
    sheet.getRange(3, 1, 1, maxCol).merge().setValue(titleText)
         .setBackground("#d35400").setFontColor("#ffffff").setFontSize(12)
         .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Row 4: SESSION DETAILS (DB0 se)
    sheet.getRange(4, 1, 1, maxCol).merge().setValue("Current Session: " + config.currentSession)
         .setBackground("#f4f6f7").setFontColor("#2d3436").setFontSize(10)
         .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Row 5: COLUMN HEADERS (UID, Name, etc.)
    sheet.getRange(5, 1, 1, maxCol).setValues([prof.cols])
         .setBackground("#2c3e50").setFontColor("#ffffff").setFontSize(10)
         .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // ==========================================
    // ⚙️ FORMATTING & SIZING
    // ==========================================
    
    // Header Borders
    sheet.getRange(1, 1, 5, maxCol).setBorder(true, true, true, true, true, true, "#bdc3c7", SpreadsheetApp.BorderStyle.SOLID);
    
    // Frozen Rows (Row 6 se data start hoga aur scroll par headers fix rahenge)
    sheet.setFrozenRows(5);
    
    // Professional Row Heights
    sheet.setRowHeight(1, 40);
    sheet.setRowHeight(2, 25);
    sheet.setRowHeight(3, 30);
    sheet.setRowHeight(4, 25);
    sheet.setRowHeight(5, 30);
    
    // Basic Column Widths set karna taaki text hide na ho
    for (var i = 1; i <= maxCol; i++) {
      sheet.setColumnWidth(i, 115); 
    }

    Logger.log("✅ Successfully generated dynamic structure for: " + prof.name);
  });
  
  return "DB8 Dynamic Setup Complete! Ready for data entry in Row 6.";
}

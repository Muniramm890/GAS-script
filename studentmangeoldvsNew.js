// ============================================================================
// 🎓 STUDENT DIRECTORY — DB2 (Pending) → DB1 (Official) PROMOTION ENGINE
// ============================================================================


function getPendingRegistrations() {
  try {
    const ss = getDB2();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { status: false, message: "Users sheet not found" };

    const data = sheet.getDataRange().getValues();
    const list = [];

    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[13] || "ACTIVE").toUpperCase(); // Col N
      const promoStatus = String(row[14] || "").toUpperCase(); // Col O

      // Sirf verified (ACTIVE) aur abhi tak promote na hui rows dikhao
      if (status === "ACTIVE" && promoStatus !== "PROMOTED") {
        
        // 🟢 DOB Format Fix (Sheet se HTML Date Input ke liye YYYY-MM-DD format banan)
        var rawDob = row[15] ? new Date(row[15]) : null;
        var formattedDob = (rawDob && !isNaN(rawDob.getTime())) ? Utilities.formatDate(rawDob, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";

        list.push({
          uid: row[0],                  // A
          name: row[1],                 // B
          father: row[2],               // C
          mobile: row[3],               // D
          email: row[4],                // E
          pin: row[6],                  // G
          photo: row[7],                // H
          state: row[8],                // I
          address: row[9],              // J
          regDate: formatDate(row[10]), // K
          // 🟢 NEW COLUMNS (P to T)
          dob: formattedDob,            // P 
          gender: row[16] || "",        // Q
          mother: row[17] || "",        // R
          classVal: row[18] || "",      // S
          section: row[19] || ""        // T
        });
      }
    }

    return { status: true, message: "Fetched", data: list };
  } catch (e) {
    return { status: false, message: "Error: " + e.toString() };
  }
}


/**
 * Ek DB2 (pending) student ko DB1 (official) Student_Data me promote karta hai.
 * Admin ko missing fields (roll, dob, gender, class, section, admDate, mother) bharni hoti hain.
 */
function promoteToOfficial(params) {
  try {
    const db2Sheet = getDB2().getSheetByName("Users");
    const db1Sheet = getDB1().getSheetByName("Student_Data");
    if (!db2Sheet || !db1Sheet) return { status: false, message: "Sheets missing" };

    const sourceUid = String(params.sourceUid || "").trim(); // DB2 wala temp UID
    const db2Data = db2Sheet.getDataRange().getValues();
    let sourceRow = -1, sourceRecord = null;

    for (var i = 1; i < db2Data.length; i++) {
      if (String(db2Data[i][0]).trim() === sourceUid) {
        sourceRow = i + 1;
        sourceRecord = db2Data[i];
        break;
      }
    }

    if (sourceRow === -1) return { status: false, message: "Source registration not found" };

    // 🔴 Duplicate Guard: same email pehle se DB1 me toh nahi
    const email = String(sourceRecord[4] || "").trim().toLowerCase();
    const db1Data = db1Sheet.getDataRange().getValues();
    for (var j = 6; j < db1Data.length; j++) { // Row 6 se data start
      if (String(db1Data[j][11] || "").trim().toLowerCase() === email) {
        return { status: false, message: "This email already exists as an Official Student!" };
      }
    }

    // Naya Official UID Generate karo
    const newUID = "STU" + Math.floor(Date.now() / 1000);

    // Password: Admin reset karna chahta hai ya purana rakhna hai
    const finalPassword = (params.resetPassword === "true" || params.resetPassword === true)
      ? (params.newPassword || "Student@123")
      : sourceRecord[5]; // Purana password (Col F)

    // DB1 Student_Data row banate hain (A se O tak — exact column order)
    const newRow = [
      newUID,                          // A: ID
      params.roll || "",               // B: Roll
      sourceRecord[1],                 // C: Name (from DB2)
      params.dob || "",                // D: DOB (admin fills)
      params.gender || "",             // E: Gender (admin fills)
      params.classVal || "",           // F: Class (admin fills)
      params.section || "",            // G: Section (admin fills)
      params.admDate ? new Date(params.admDate) : new Date(), // H: Adm Date
      sourceRecord[2],                 // I: Father (from DB2)
      params.mother || "",             // J: Mother (admin fills)
      sourceRecord[3],                 // K: Mobile (from DB2)
      sourceRecord[4],                 // L: Email (from DB2)
      sourceRecord[9],                 // M: Address (from DB2)
      finalPassword,                   // N: Password
      sourceRecord[7]                  // O: Photo (from DB2)
    ];

    db1Sheet.appendRow(newRow);

    // DB2 row ko PROMOTED mark karo (delete nahi, audit trail ke liye)
    db2Sheet.getRange(sourceRow, 15).setValue("PROMOTED"); // Col O

    sendPushNotification(
      "Student Promoted to Official",
      `${sourceRecord[1]} has been officially admitted with ID ${newUID}.`,
      { type: "INFO", uid: newUID }
    );

    return { status: true, message: "Student Promoted Successfully!", data: { newUid: newUID } };

  } catch (e) {
    return { status: false, message: "Promotion Error: " + e.toString() };
  }
}

/**
 * Official (DB1) student ko permanently delete karta hai.
 */
function deleteOfficialStudent(uid) {
  try {
    const sheet = getDB1().getSheetByName("Student_Data");
    const data = sheet.getDataRange().getValues();
    const targetUid = String(uid).trim();

    for (var i = 5; i < data.length; i++) { // Row 6 se data start (index 5)
      if (String(data[i][0]).trim() === targetUid) {
        sheet.deleteRow(i + 1);
        return { status: true, message: "Official student record deleted" };
      }
    }
    return { status: false, message: "Student not found" };
  } catch (e) {
    return { status: false, message: "Error: " + e.toString() };
  }
}

/**
 * Pending (DB2) registration ko reject/delete karta hai.
 */
function deletePendingRegistration(uid) {
  try {
    const sheet = getDB2().getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const targetUid = String(uid).trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetUid) {
        sheet.deleteRow(i + 1);
        return { status: true, message: "Registration rejected & removed" };
      }
    }
    return { status: false, message: "Record not found" };
  } catch (e) {
    return { status: false, message: "Error: " + e.toString() };
  }
}

/**
 * 🔑 UNIVERSAL PASSWORD RESET — Official (DB1) ya Pending (DB2) dono ke liye kaam karega.
 */
function resetStudentPassword(uid, source, newPassword) {
  try {
    const targetUid = String(uid).trim();
    const finalPass = String(newPassword || "Student@123").trim();

    if (source === "OFFICIAL") {
      const sheet = getDB1().getSheetByName("Student_Data");
      const data = sheet.getDataRange().getValues();
      for (var i = 5; i < data.length; i++) {
        if (String(data[i][0]).trim() === targetUid) {
          sheet.getRange(i + 1, 14).setValue(finalPass); // Col N = Password
          return { status: true, message: "Password reset successfully (Official)" };
        }
      }
    } else if (source === "PENDING") {
      const sheet = getDB2().getSheetByName("Users");
      const data = sheet.getDataRange().getValues();
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][0]).trim() === targetUid) {
          sheet.getRange(j + 1, 6).setValue(finalPass); // Col F = Password
          return { status: true, message: "Password reset successfully (Pending)" };
        }
      }
    }
    return { status: false, message: "Record not found for password reset" };
  } catch (e) {
    return { status: false, message: "Error: " + e.toString() };
  }
}
// ============================================================================
// 🎓 STUDENT DIRECTORY — FETCH ALL OFFICIAL STUDENTS (DB1)
// ============================================================================
function getAllOfficialStudents() {
  try {
    var ss = getDB1();
    var sheet = ss.getSheetByName("Student_Data");
    
    if (!sheet) {
      return { status: false, message: "Student_Data sheet not found in DB1!" };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 6) return { status: true, data: [] };

    // Fetching exactly 15 columns (A to O)
    var data = sheet.getRange(6, 1, lastRow - 5, 15).getValues();
    var studentsList = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var uid = String(row[0] || "").trim();
      
      // Skip empty rows
      if (uid !== "") {
        studentsList.push({
          uid: uid,
          roll: String(row[1] || "").trim(),
          name: String(row[2] || "").trim(),
          dob: row[3] ? formatDate(row[3]) : "",           // Safe Date Check
          gender: String(row[4] || "").trim(),
          classVal: String(row[5] || "").trim(),
          section: String(row[6] || "").trim(),
          admDate: row[7] ? formatDate(row[7]) : "",       // Safe Date Check
          father: String(row[8] || "").trim(),
          mother: String(row[9] || "").trim(),
          mobile: String(row[10] || "").trim(),
          email: String(row[11] || "").trim(),
          address: String(row[12] || "").trim(),
          // row[13] is Password (Skipped for Security)
          photo: String(row[14] || "").trim()
        });
      }
    }

    return { status: true, success: true, message: "Official students fetched", data: studentsList };

  } catch (e) {
    return { status: false, message: "Error fetching official students: " + e.toString() };
  }
}

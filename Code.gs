// ============================================================
// FSIS Logger — Google Apps Script Backend (Code.gs)
// ============================================================
// Paste this entire file into: Google Sheet > Extensions > Apps Script
// Then deploy as: Web App | Execute as: Me | Who has access: Anyone
// ============================================================

// ── CONFIGURATION ───────────────────────────────────────────
// 1. Replace with your Google Drive folder ID for photo storage
//    (Open the folder in Drive and copy the ID from the URL)
var DRIVE_FOLDER_ID = "1dZPGdfM8hKxN8LzrP_XrkxD2-hs9LYZA";

// 2. Replace with a strong secret used by the admin panel
var ADMIN_SECRET = "YOUR_ADMIN_SECRET_HERE";
// ────────────────────────────────────────────────────────────

/**
 * Entry point for all HTTP POST requests from the frontend.
 * All requests send JSON in the POST body with an "action" field.
 */
function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === "ping")         result = { ok: true };
    else if (action === "login")   result = handleLogin(body);
    else if (action === "list_users")   result = handleListUsers(body);
    else if (action === "create_user")  result = handleCreateUser(body);
    else if (action === "read")    result = handleRead(body);
    else if (action === "insert")  result = handleInsert(body);
    else if (action === "update")  result = handleUpdate(body);
    else if (action === "delete")  result = handleDelete(body);
    else if (action === "upload")  result = handleUpload(body);
    else if (action === "patch_photo_url") result = handlePatchPhotoUrl(body);
    else if (action === "patch_lat_lng") result = handlePatchLatLng(body);
    else result = { error: "Unknown action: " + action };

  } catch (err) {
    result = { error: err.message || String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Health-check for GET requests */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: "FSIS Logger API" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ─────────────────────────────────────────────────

/** Get a sheet tab by name, throw if not found */
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("Sheet tab not found: " + name);
  return sheet;
}

/** Read all rows from a sheet tab and return as array of objects */
function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // Convert Date objects to ISO strings
      if (val instanceof Date) {
        obj[headers[j]] = val.toISOString();
      } else {
        obj[headers[j]] = val === "" ? null : val;
      }
    }
    rows.push(obj);
  }
  return rows;
}

/** Generate a UUID v4 */
function generateUUID() {
  return Utilities.getUuid();
}

/** Find the row number (1-indexed) of a record by its id column value */
function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return -1;
  var headers = data[0];
  var idCol = headers.indexOf("id");
  if (idCol === -1) return -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

// ── ACTION HANDLERS ─────────────────────────────────────────

/**
 * LOGIN — checks username + password against the users tab
 * Body: { username, password }
 * Returns: { data: { id, username, display_name, role } } or { error }
 */
function handleLogin(body) {
  var username = (body.username || "").trim().toLowerCase();
  var password = body.password || "";
  if (!username || !password) return { error: "Username and password required." };

  var sheet = getSheet("users");
  var users = sheetToObjects(sheet);
  var user = users.find(function(u) {
    return String(u.username || "").trim().toLowerCase() === username
        && String(u.password || "") === String(password);
  });

  if (!user) return { error: "Invalid username or password." };

  return {
    data: [{
      id: user.id,
      username: user.username,
      display_name: user.display_name || user.username,
      role: user.role || "user"
    }]
  };
}

/**
 * LIST USERS — returns all users (admin only)
 * Body: { adminSecret }
 */
function handleListUsers(body) {
  if ((body.adminSecret || "") !== ADMIN_SECRET) return { error: "Invalid admin secret." };
  var sheet = getSheet("users");
  var users = sheetToObjects(sheet).map(function(u) {
    return { id: u.id, username: u.username, display_name: u.display_name, role: u.role, created_at: u.created_at };
  });
  return { data: users };
}

/**
 * CREATE USER — appends a new user row (admin only)
 * Body: { adminSecret, username, displayName, password, role }
 */
function handleCreateUser(body) {
  if ((body.adminSecret || "") !== ADMIN_SECRET) return { error: "Invalid admin secret." };

  var username = (body.username || "").trim();
  var displayName = (body.displayName || "").trim();
  var password = body.password || "";
  var role = body.role || "user";

  if (!username) return { error: "Username is required." };
  if (!password || password.length < 4) return { error: "Password must be at least 4 characters." };

  var sheet = getSheet("users");
  var existingUsers = sheetToObjects(sheet);
  var duplicate = existingUsers.find(function(u) {
    return String(u.username || "").toLowerCase() === username.toLowerCase();
  });
  if (duplicate) return { error: "Username already exists." };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = headers.map(function(h) {
    if (h === "id") return generateUUID();
    if (h === "username") return username;
    if (h === "display_name") return displayName;
    if (h === "password") return password;
    if (h === "role") return role;
    if (h === "created_at") return new Date().toISOString();
    return "";
  });
  sheet.appendRow(newRow);
  return { data: { username: username } };
}

/**
 * READ — returns all rows of a table tab
 * Body: { table }
 */
function handleRead(body) {
  var table = body.table || "";
  if (!table) return { error: "table name required." };
  var sheet = getSheet(table);
  return { data: sheetToObjects(sheet) };
}

/**
 * INSERT — appends a row to a table tab
 * Body: { table, row: { field: value, ... } }
 */
function handleInsert(body) {
  var table = body.table || "";
  var row = body.row || {};
  if (!table) return { error: "table required." };

  var sheet = getSheet(table);
  if (table === "fire_drill_logbook") {
    ensureColumnExists(sheet, "owner_name");
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var id = generateUUID();
  var createdAt = new Date().toISOString();

  var newRow = headers.map(function(h) {
    if (h === "id") return id;
    if (h === "created_at") return createdAt;
    var v = row[h];
    return (v === undefined || v === null) ? "" : v;
  });

  sheet.appendRow(newRow);
  return { data: { id: id, created_at: createdAt } };
}

/**
 * UPDATE — find row by id and update specified columns
 * Body: { table, id, row: { field: value, ... } }
 */
function handleUpdate(body) {
  var table = body.table || "";
  var id = body.id || "";
  var updates = body.row || {};
  if (!table || !id) return { error: "table and id required." };

  var sheet = getSheet(table);
  if (table === "fire_drill_logbook") {
    ensureColumnExists(sheet, "owner_name");
  }
  var rowNum = findRowById(sheet, id);
  if (rowNum < 0) return { error: "Record not found: " + id };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

  var newRow = headers.map(function(h, idx) {
    if (h === "id") return currentRow[idx]; // never overwrite id
    if (h === "created_at") return currentRow[idx]; // never overwrite created_at
    if (updates.hasOwnProperty(h) && updates[h] !== undefined && updates[h] !== null) {
      return updates[h];
    }
    return currentRow[idx];
  });

  sheet.getRange(rowNum, 1, 1, headers.length).setValues([newRow]);
  return { data: { id: id } };
}

function ensureColumnExists(sheet, columnName) {
  if (!sheet || !columnName) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(columnName) !== -1) return;
  sheet.getRange(1, lastCol + 1).setValue(columnName);
}

/**
 * DELETE — find row by id and delete it
 * Body: { table, id }
 */
function handleDelete(body) {
  var table = body.table || "";
  var id = body.id || "";
  if (!table || !id) return { error: "table and id required." };

  var sheet = getSheet(table);
  var rowNum = findRowById(sheet, id);
  if (rowNum < 0) return { error: "Record not found: " + id };

  sheet.deleteRow(rowNum);
  return { data: { id: id } };
}

/**
 * ─────────────────────────────────────────────────────────────────────────
 * RUN THIS ONCE from the Apps Script editor to add photo_url columns
 * ─────────────────────────────────────────────────────────────────────────
 * Select "setupPhotoUrlColumns" from the dropdown → click ▶ Run
 */
function setupPhotoUrlColumns() {
  var tables = ["inspection_logbook", "occupancy_logbook", "fire_drill_logbook"];
  var columnsToAdd = ["photo_url", "photo_taken_at", "latitude", "longitude"];

  tables.forEach(function(tableName) {
    var sheet;
    try { sheet = getSheet(tableName); }
    catch(e) { Logger.log("⚠️ Sheet not found: " + tableName); return; }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("📋 " + tableName + " columns: " + headers.join(", "));

    columnsToAdd.forEach(function(col) {
      if (headers.indexOf(col) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
        Logger.log("✅ Added '" + col + "' to " + tableName);
      } else {
        Logger.log("✔️  '" + col + "' already exists in " + tableName);
      }
    });
  });

  Logger.log("Done — check your Google Sheet.");
}

/**
 * RUN ONCE from the Apps Script editor to create the Fire Drill logbook tab
 * with column headers matching `fire_drill_certificate.html` placeholders
 * (&lt;control_number&gt;, &lt;DATE&gt;, &lt;Building Name&gt;, &lt;ADDRESS&gt;, etc.).
 * Does not overwrite row 1 if it already has data.
 */
function setupFireDrillLogbookSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = "fire_drill_logbook";
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  var headers = [
    "id",
    "created_at",
    "control_number",
    "certificate_date",
    "building_name",
    "owner_name",
    "address",
    "day_issued",
    "month_year_issued",
    "date_valid",
    "amount_paid",
    "or_number",
    "date_paid",
  ];

  var a1 = sheet.getRange(1, 1).getValue();
  if (a1 === "" || a1 === null) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log("✅ Wrote header row on " + name);
  } else {
    Logger.log("⚠️ Row 1 already has data — left unchanged. Headers expected: " + headers.join(", "));
  }

  Logger.log("Fire Drill logbook tab ready. Run setupPhotoUrlColumns if you need photo/location columns.");
}

/**
 * ONE-CLICK SETUP for Google Sheets (run from the Apps Script editor)
 * ─────────────────────────────────────────────────────────────────────────
 * 1. Open the spreadsheet bound to this script (or the spreadsheet
 *    linked to this project).
 * 2. Extensions → Apps Script → select "addFireDrillLogbookToSpreadsheet"
 *    in the function dropdown → click Run ▶
 * 3. Authorize if prompted.
 *
 * Creates the tab "fire_drill_logbook" with all certificate columns, then
 * adds photo_url, photo_taken_at, latitude, longitude if missing (same as
 * other logbooks). Safe to run more than once.
 */
function addFireDrillLogbookToSpreadsheet() {
  setupFireDrillLogbookSheet();

  var tableName = "fire_drill_logbook";
  var sheet;
  try {
    sheet = getSheet(tableName);
  } catch (e) {
    Logger.log("❌ " + e.message);
    return;
  }

  var columnsToAdd = ["owner_name", "photo_url", "photo_taken_at", "latitude", "longitude"];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  columnsToAdd.forEach(function (col) {
    if (headers.indexOf(col) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
      Logger.log("✅ Added '" + col + "' to " + tableName);
    } else {
      Logger.log("✔️  '" + col + "' already exists in " + tableName);
    }
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  });

  Logger.log("Done — fire_drill_logbook is ready for the web app.");
}

/**
 * PATCH PHOTO URL — find existing row by id, then find (or create) the
 * photo_url column and write the Drive URL directly to that cell.
 * Body: { table, id, url }
 * Returns: { data: { id } } or { error }
 */
function handlePatchPhotoUrl(body) {
  var table = body.table || "";
  var id = body.id || "";
  var url = body.url || "";
  if (!table || !id) return { error: "table and id required." };
  if (!url) return { error: "url required." };

  var sheet = getSheet(table);
  var rowNum = findRowById(sheet, id);
  if (rowNum < 0) return { error: "Record not found: " + id };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIdx = headers.indexOf("photo_url");

  // If the column doesn't exist yet, add it at the end
  if (colIdx === -1) {
    var newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue("photo_url");
    colIdx = newCol - 1; // 0-indexed
  }

  // Write the URL directly to the correct cell (sheet is 1-indexed)
  sheet.getRange(rowNum, colIdx + 1).setValue(url);
  return { data: { id: id } };
}

/**
 * PATCH LAT/LNG — write latitude & longitude cells by row id (creates columns if missing).
 * Body: { table, id, latitude, longitude }
 */
function handlePatchLatLng(body) {
  var table = body.table || "";
  var id = body.id || "";
  var lat = body.latitude;
  var lng = body.longitude;
  if (!table || !id) return { error: "table and id required." };
  if (lat === undefined || lat === null || lng === undefined || lng === null) {
    return { error: "latitude and longitude required." };
  }

  var sheet = getSheet(table);
  var rowNum = findRowById(sheet, id);
  if (rowNum < 0) return { error: "Record not found: " + id };

  var columnsToAdd = ["latitude", "longitude"];
  columnsToAdd.forEach(function (col) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf(col) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
    }
  });

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var latCol = headers.indexOf("latitude");
  var lngCol = headers.indexOf("longitude");
  if (latCol === -1 || lngCol === -1) {
    return { error: "Could not resolve latitude/longitude columns." };
  }

  var latNum = Number(lat);
  var lngNum = Number(lng);
  if (isNaN(latNum) || isNaN(lngNum)) {
    return { error: "Invalid latitude or longitude." };
  }

  sheet.getRange(rowNum, latCol + 1).setValue(latNum);
  sheet.getRange(rowNum, lngCol + 1).setValue(lngNum);
  return { data: { id: id } };
}

/**
 * UPLOAD — decode base64 file, save to Drive, return shareable URL
 * Body: { filename, mimeType, base64Data }
 * Returns: { data: { url } }
 */
function handleUpload(body) {
  var filename = body.filename || ("upload-" + Date.now());
  var mimeType = body.mimeType || "image/jpeg";
  var base64Data = body.base64Data || "";

  if (!base64Data) return { error: "No file data provided." };
  if (!DRIVE_FOLDER_ID || DRIVE_FOLDER_ID === "YOUR_DRIVE_FOLDER_ID_HERE") {
    return {
      error: "Google Drive Folder ID not configured. Please open Code.gs in Apps Script and set DRIVE_FOLDER_ID."
    };
  }

  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
    var file = folder.createFile(blob);

    // Try to make the file publicly viewable so the URL renders in the browser
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      // Ignore "Tinanggihang bigyan ng access: DriveApp" errors caused by 
      // strict Google Workspace organization policies. The file is already uploaded.
      // (The user can manually share the 'FSIS Storage' folder instead).
    }

    var fileId = file.getId();
    var url = "https://drive.google.com/uc?export=view&id=" + fileId;

    return { data: { url: url, fileId: fileId } };
  } catch (err) {
    return { error: "Drive upload failed: " + err.message };
  }
}

/**
 * ─────────────────────────────────────────────────────────────────────────
 * RUN THIS FUNCTION ONCE FROM THE APPS SCRIPT EDITOR BEFORE DEPLOYING
 * ─────────────────────────────────────────────────────────────────────────
 * 1. Open this file in the Apps Script editor (Extensions → Apps Script)
 * 2. Select "testDriveAccess" from the function dropdown at the top
 * 3. Click ▶ Run — it will ask you to authorize Google Drive access
 * 4. Accept the permission prompt
 * 5. Then: Deploy → New Deployment (Web App, Execute as Me, Anyone)
 * ─────────────────────────────────────────────────────────────────────────
 */
function testDriveAccess() {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log("✅ Drive access OK. Folder name: " + folder.getName());
    Logger.log("Folder ID: " + DRIVE_FOLDER_ID);
  } catch (err) {
    Logger.log("❌ Drive access FAILED: " + err.message);
    Logger.log("Check that DRIVE_FOLDER_ID is correct and you have Editor access to the folder.");
  }
}

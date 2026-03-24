// ============================================================
// FSIS Logger — Google Apps Script Backend (Code.gs)
// ============================================================
// Paste this entire file into: Google Sheet > Extensions > Apps Script
// Then deploy as: Web App | Execute as: Me | Who has access: Anyone
// ============================================================

// ── CONFIGURATION ───────────────────────────────────────────
// 1. Replace with your Google Drive folder ID for photo storage
//    (Open the folder in Drive and copy the ID from the URL)
var DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID_HERE";

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
  var rowNum = findRowById(sheet, id);
  if (rowNum < 0) return { error: "Record not found: " + id };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

  var newRow = headers.map(function(h, idx) {
    if (h === "id") return currentRow[idx]; // never overwrite id
    if (updates.hasOwnProperty(h) && updates[h] !== undefined && updates[h] !== null) {
      return updates[h];
    }
    return currentRow[idx];
  });

  sheet.getRange(rowNum, 1, 1, headers.length).setValues([newRow]);
  return { data: { id: id } };
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
 * UPLOAD — decode base64 file, save to Drive, return shareable URL
 * Body: { filename, mimeType, base64Data }
 * Returns: { data: { url } }
 */
function handleUpload(body) {
  var filename = body.filename || ("upload-" + Date.now());
  var mimeType = body.mimeType || "image/jpeg";
  var base64Data = body.base64Data || "";

  if (!base64Data) return { error: "No file data provided." };
  if (DRIVE_FOLDER_ID === "YOUR_DRIVE_FOLDER_ID_HERE") {
    return { error: "DRIVE_FOLDER_ID is not configured in Code.gs." };
  }

  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var file = folder.createFile(blob);

  // Make the file publicly viewable so the URL renders in the browser
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Use the direct image URL format (works for <img> tags)
  var fileId = file.getId();
  var url = "https://drive.google.com/uc?export=view&id=" + fileId;

  return { data: { url: url } };
}

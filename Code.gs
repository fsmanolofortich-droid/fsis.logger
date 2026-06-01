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

// 2. Admin access uses the "users" tab: set role to "admin" on an account row.
//    Sign in on the login page with that username/password to open the admin panel.
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
    else if (action === "update_user")  result = handleUpdateUser(body);
    else if (action === "suspend_user") result = handleSuspendUser(body);
    else if (action === "read")    result = handleRead(body);
    else if (action === "insert")  result = handleInsert(body);
    else if (action === "update")  result = handleUpdate(body);
    else if (action === "delete")  result = handleDelete(body);
    else if (action === "upload")  result = handleUpload(body);
    else if (action === "patch_photo_url") result = handlePatchPhotoUrl(body);
    else if (action === "patch_lat_lng") result = handlePatchLatLng(body);
    else if (action === "patch_expires_on") result = handlePatchExpiresOn(body);
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
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  // Use full header width (not getDataRange) so new columns like expires_on are never skipped.
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  var rows = [];
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "GMT";
  for (var i = 0; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      var key = String(headers[j] || "").trim();
      // Convert Date objects to ISO strings
      if (val instanceof Date) {
        // IMPORTANT: Using toISOString() converts to UTC and can shift the date (off-by-1/-2 days)
        // depending on timezone. Most of our sheet dates are intended as calendar dates, so we
        // return a YYYY-MM-DD string in the spreadsheet timezone.
        obj[key] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
      } else {
        obj[key] = val === "" ? null : val;
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

// ── USER ACCOUNT HELPERS ────────────────────────────────────

function parseSuspendedFlag(value) {
  if (value === true || value === 1) return true;
  var s = String(value || "").trim().toLowerCase();
  return s === "true" || s === "yes" || s === "1" || s === "suspended";
}

function isUserSuspended(user) {
  return parseSuspendedFlag(user && user.suspended);
}

function findUserByLogin(sheet, loginId) {
  var users = sheetToObjects(sheet);
  var key = String(loginId || "").trim().toLowerCase();
  return users.find(function (u) {
    var uName = String(u.username || "").trim().toLowerCase();
    var uEmail = String(u.email || "").trim().toLowerCase();
    return uName === key || (uEmail && uEmail === key);
  }) || null;
}

// ── ACTION HANDLERS ─────────────────────────────────────────

/**
 * LOGIN — checks username + password against the users tab
 * Body: { username, password }
 * Returns: { data: { id, username, display_name, role } } or { error }
 */
function handleLogin(body) {
  var loginId = (body.username || body.email || "").trim().toLowerCase();
  var password = body.password || "";
  if (!loginId || !password) return { error: "Username and password required." };

  var sheet = getSheet("users");
  ensureColumnExists(sheet, "email");
  ensureColumnExists(sheet, "suspended");
  var user = findUserByLogin(sheet, loginId);
  if (!user || String(user.password || "") !== String(password)) {
    return { error: "Invalid username or password." };
  }
  if (isUserSuspended(user)) {
    return { error: "This account has been suspended. Contact your administrator." };
  }

  return {
    data: [{
      id: user.id,
      username: user.username,
      display_name: user.display_name || user.username,
      role: user.role || "user"
    }]
  };
}

/** Verify username/password and require role === admin. Returns error object or null. */
function requireAdminUser(body) {
  var username = (body.username || "").trim().toLowerCase();
  var password = body.password || "";
  if (!username || !password) return { error: "Admin sign-in required." };

  var sheet = getSheet("users");
  ensureColumnExists(sheet, "suspended");
  var user = findUserByLogin(sheet, username);
  if (!user || String(user.password || "") !== String(password)) {
    return { error: "Invalid admin credentials." };
  }
  if (isUserSuspended(user)) {
    return { error: "This admin account is suspended." };
  }
  if (String(user.role || "user").toLowerCase() !== "admin") {
    return { error: "Admin access required." };
  }
  return null;
}

/**
 * LIST USERS — returns all users (admin only)
 * Body: { username, password }
 */
function handleListUsers(body) {
  var authErr = requireAdminUser(body);
  if (authErr) return authErr;
  var sheet = getSheet("users");
  ensureColumnExists(sheet, "email");
  ensureColumnExists(sheet, "suspended");
  var users = sheetToObjects(sheet).map(function(u) {
    return {
      id: u.id,
      username: u.username,
      email: u.email || "",
      display_name: u.display_name,
      role: u.role,
      suspended: isUserSuspended(u),
      created_at: u.created_at
    };
  });
  return { data: users };
}

/**
 * CREATE USER — appends a new user row (admin only)
 * Body: { username, password, newUsername, displayName, newPassword, role }
 *   username/password = signed-in admin credentials
 *   newUsername, newPassword, displayName, role = account to create
 */
function handleCreateUser(body) {
  var authErr = requireAdminUser(body);
  if (authErr) return authErr;

  var newUsername = (body.newUsername || "").trim();
  var newEmail = (body.email || "").trim();
  var displayName = (body.displayName || "").trim();
  var newPassword = body.newPassword || "";
  var role = body.role || "user";

  if (!newUsername) return { error: "Username is required." };
  if (!newPassword || newPassword.length < 4) return { error: "Password must be at least 4 characters." };

  var sheet = getSheet("users");
  ensureColumnExists(sheet, "email");
  ensureColumnExists(sheet, "suspended");
  var existingUsers = sheetToObjects(sheet);
  var duplicate = existingUsers.find(function(u) {
    return String(u.username || "").toLowerCase() === newUsername.toLowerCase();
  });
  if (duplicate) return { error: "Username already exists." };
  if (newEmail) {
    var dupEmail = existingUsers.find(function (u) {
      return String(u.email || "").trim().toLowerCase() === newEmail.toLowerCase();
    });
    if (dupEmail) return { error: "Email already in use." };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = headers.map(function(h) {
    if (h === "id") return generateUUID();
    if (h === "username") return newUsername;
    if (h === "email") return newEmail;
    if (h === "display_name") return displayName;
    if (h === "password") return newPassword;
    if (h === "role") return role;
    if (h === "suspended") return false;
    if (h === "created_at") return new Date().toISOString();
    return "";
  });
  sheet.appendRow(newRow);
  return { data: { username: newUsername, email: newEmail } };
}

/**
 * SUSPEND USER — suspend or restore an account (admin only)
 * Body: { username, password, userId, suspended: true|false }
 */
function handleSuspendUser(body) {
  var authErr = requireAdminUser(body);
  if (authErr) return authErr;

  var userId = body.userId || body.id;
  if (!userId) return { error: "User id required." };

  if (body.suspended === undefined || body.suspended === null) {
    return { error: "suspended flag required (true or false)." };
  }
  var suspendFlag = body.suspended === true
    || String(body.suspended).toLowerCase() === "true";

  var sheet = getSheet("users");
  ensureColumnExists(sheet, "suspended");
  var rowNum = findRowById(sheet, userId);
  if (rowNum < 0) return { error: "User not found." };

  var allUsers = sheetToObjects(sheet);
  var target = allUsers.find(function (u) { return String(u.id) === String(userId); });
  if (!target) return { error: "User not found." };

  var adminLogin = String(body.username || "").trim().toLowerCase();
  var adminUser = findUserByLogin(sheet, adminLogin);
  if (adminUser && String(adminUser.id) === String(userId) && suspendFlag) {
    return { error: "You cannot suspend your own account." };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var suspCol = headers.indexOf("suspended") + 1;
  if (suspCol < 1) return { error: "suspended column missing." };

  sheet.getRange(rowNum, suspCol).setValue(suspendFlag ? "TRUE" : "FALSE");

  return {
    data: {
      id: userId,
      username: target.username,
      suspended: suspendFlag
    }
  };
}

/**
 * UPDATE USER — edit account (admin only)
 * Body: { username, password, userId, username?, email?, displayName?, password?, role? }
 *   username/password at top level = admin credentials
 *   userId = row to edit; other fields = updates (password optional)
 */
function handleUpdateUser(body) {
  var authErr = requireAdminUser(body);
  if (authErr) return authErr;

  var userId = body.userId || body.id;
  if (!userId) return { error: "User id required." };

  var sheet = getSheet("users");
  ensureColumnExists(sheet, "email");
  var rowNum = findRowById(sheet, userId);
  if (rowNum < 0) return { error: "User not found." };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var allUsers = sheetToObjects(sheet);
  var current = allUsers.find(function (u) { return String(u.id) === String(userId); });
  if (!current) return { error: "User not found." };

  var targetUsername = body.targetUsername != null
    ? String(body.targetUsername).trim()
    : String(current.username || "").trim();
  var targetEmail = body.email != null
    ? String(body.email).trim()
    : String(current.email || "").trim();
  var targetDisplay = body.displayName != null
    ? String(body.displayName).trim()
    : String(current.display_name || "").trim();
  var targetRole = body.role != null
    ? String(body.role).trim()
    : String(current.role || "user");
  var newPassword = body.newPassword != null ? String(body.newPassword) : "";

  if (!targetUsername) return { error: "Username is required." };
  if (newPassword && newPassword.length < 4) {
    return { error: "Password must be at least 4 characters." };
  }

  var dupUser = allUsers.find(function (u) {
    return String(u.id) !== String(userId)
      && String(u.username || "").trim().toLowerCase() === targetUsername.toLowerCase();
  });
  if (dupUser) return { error: "Username already exists." };

  if (targetEmail) {
    var dupEmail = allUsers.find(function (u) {
      return String(u.id) !== String(userId)
        && String(u.email || "").trim().toLowerCase() === targetEmail.toLowerCase();
    });
    if (dupEmail) return { error: "Email already in use." };
  }

  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || "").trim();
    var col = c + 1;
    if (h === "username") sheet.getRange(rowNum, col).setValue(targetUsername);
    else if (h === "email") sheet.getRange(rowNum, col).setValue(targetEmail);
    else if (h === "display_name") sheet.getRange(rowNum, col).setValue(targetDisplay);
    else if (h === "role") sheet.getRange(rowNum, col).setValue(targetRole);
    else if (h === "password" && newPassword) sheet.getRange(rowNum, col).setValue(newPassword);
  }

  return {
    data: {
      id: userId,
      username: targetUsername,
      email: targetEmail,
      display_name: targetDisplay,
      role: targetRole
    }
  };
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
  if (table === "inspection_logbook" || table === "occupancy_logbook") {
    ensureColumnExists(sheet, "io_remarks");
  }
  if (table === "inspection_logbook") {
    ensureColumnExists(sheet, "expires_on");
  }
  if (table === "fsec_building_plan_logbook") {
    ensureColumnExists(sheet, "fsec_number");
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var id = generateUUID();
  var createdAt = new Date().toISOString();

  var newRow = headers.map(function(h) {
    var key = String(h || "").trim();
    if (key === "id") return id;
    if (key === "created_at") return createdAt;
    // Allow sheet headers with accidental trailing/leading spaces.
    var v = row[key];
    if ((v === undefined || v === null || v === "") && row.expires_on) {
      if (key === "expires_on" || key === "Expires On" || key === "expire_date" || key === "fsic_valid_until") {
        v = row.expires_on;
      }
    }
    if ((v === undefined || v === null || v === "") && row.fsic_valid_until && key === "fsic_valid_until") {
      v = row.fsic_valid_until;
    }
    return (v === undefined || v === null) ? "" : v;
  });

  sheet.appendRow(newRow);
  if (table === "inspection_logbook" && (row.expires_on || row.fsic_valid_until)) {
    writeExpiresOnForRow(sheet, id, row.expires_on || row.fsic_valid_until);
  }
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
  if (table === "inspection_logbook" || table === "occupancy_logbook") {
    ensureColumnExists(sheet, "io_remarks");
  }
  if (table === "inspection_logbook") {
    ensureColumnExists(sheet, "expires_on");
  }
  if (table === "fsec_building_plan_logbook") {
    ensureColumnExists(sheet, "fsec_number");
  }
  var rowNum = findRowById(sheet, id);
  if (rowNum < 0) return { error: "Record not found: " + id };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

  var newRow = headers.map(function(h, idx) {
    var key = String(h || "").trim();
    if (key === "id") return currentRow[idx]; // never overwrite id
    if (key === "created_at") return currentRow[idx]; // never overwrite created_at
    // Allow sheet headers with accidental trailing/leading spaces.
    if (updates.hasOwnProperty(key) && updates[key] !== undefined && updates[key] !== null) {
      return updates[key];
    }
    if (updates.expires_on && (key === "expires_on" || key === "Expires On" || key === "expire_date" || key === "fsic_valid_until")) {
      return updates.expires_on;
    }
    if (updates.fsic_valid_until && key === "fsic_valid_until") {
      return updates.fsic_valid_until;
    }
    return currentRow[idx];
  });

  sheet.getRange(rowNum, 1, 1, headers.length).setValues([newRow]);
  if (table === "inspection_logbook" && (updates.expires_on || updates.fsic_valid_until)) {
    writeExpiresOnForRow(sheet, id, updates.expires_on || updates.fsic_valid_until);
  }
  return { data: { id: id } };
}

function writeExpiresOnForRow(sheet, rowId, expiresOn) {
  if (!sheet || !rowId || !expiresOn) return false;
  var rowNum = findRowById(sheet, rowId);
  if (rowNum < 0) return false;
  var value = String(expiresOn).trim();
  var wrote = false;
  ensureColumnExists(sheet, "expires_on");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIdx = findHeaderIndex(headers, "expires_on");
  if (colIdx >= 0) {
    sheet.getRange(rowNum, colIdx + 1).setValue(value);
    wrote = true;
  }
  // Fallback: many FSIS sheets already have fsic_valid_until from clearance workflow.
  colIdx = findHeaderIndex(headers, "fsic_valid_until");
  if (colIdx >= 0) {
    sheet.getRange(rowNum, colIdx + 1).setValue(value);
    wrote = true;
  }
  return wrote;
}

function findHeaderIndex(headers, columnName) {
  var target = String(columnName || "").trim();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === target) return i;
  }
  return -1;
}

function ensureColumnExists(sheet, columnName) {
  if (!sheet || !columnName) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (findHeaderIndex(headers, columnName) !== -1) return;
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

    if (tableName === "inspection_logbook") {
      ensureColumnExists(sheet, "expires_on");
    }

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
 * RUN ONCE to add FSEC No. column on FSEC logbook table.
 * Function name to run in Apps Script: addFsecNumberColumn
 */
function addFsecNumberColumn() {
  var tableName = "fsec_building_plan_logbook";
  var sheet;
  try {
    sheet = getSheet(tableName);
  } catch (e) {
    Logger.log("❌ " + e.message);
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("fsec_number") === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("fsec_number");
    Logger.log("✅ Added 'fsec_number' to " + tableName);
  } else {
    Logger.log("✔️  'fsec_number' already exists in " + tableName);
  }
}

/**
 * RUN ONCE to add Expires On column on Inspection logbook.
 * Function name to run in Apps Script: addInspectionExpiresOnColumn
 */
function addInspectionExpiresOnColumn() {
  var tableName = "inspection_logbook";
  var sheet;
  try {
    sheet = getSheet(tableName);
  } catch (e) {
    Logger.log("❌ " + e.message);
    return;
  }

  ensureColumnExists(sheet, "expires_on");
  Logger.log("✅ Column 'expires_on' is ready on " + tableName);
}

/**
 * PATCH EXPIRES ON — write expires_on cell by row id (creates column if missing).
 * Body: { table, id, expires_on }  (YYYY-MM-DD)
 */
function handlePatchExpiresOn(body) {
  var table = body.table || "";
  var id = body.id || "";
  var expiresOn = body.expires_on;
  if (!table || !id) return { error: "table and id required." };
  if (!expiresOn) return { error: "expires_on required." };

  var sheet = getSheet(table);
  if (table === "inspection_logbook") {
    ensureColumnExists(sheet, "expires_on");
  }

  var rowNum = findRowById(sheet, id);
  if (rowNum < 0) return { error: "Record not found: " + id };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIdx = findHeaderIndex(headers, "expires_on");
  if (colIdx === -1) {
    ensureColumnExists(sheet, "expires_on");
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    colIdx = findHeaderIndex(headers, "expires_on");
  }
  if (colIdx === -1) return { error: "Could not resolve expires_on column." };

  sheet.getRange(rowNum, colIdx + 1).setValue(String(expiresOn).trim());
  return { data: { id: id, expires_on: String(expiresOn).trim() } };
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

/** Find sheet row (1-indexed) by username column. */
function findRowByUsername(sheet, username) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return -1;
  var headers = data[0];
  var userCol = headers.indexOf("username");
  if (userCol === -1) return -1;
  var target = String(username || "").trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][userCol]).trim().toLowerCase() === target) return i + 1;
  }
  return -1;
}

/** Ensure the users tab exists with standard headers. */
function ensureUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("users");
  var headers = ["id", "username", "email", "display_name", "password", "role", "suspended", "created_at"];
  if (!sheet) {
    sheet = ss.insertSheet("users");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

/**
 * Create or update an admin (or any) user row on the users tab.
 * @param {string} username
 * @param {string} password
 * @param {string} displayName
 * @param {string} role  e.g. "admin" or "user"
 */
function ensureAdminUser(username, password, displayName, role) {
  var sheet = ensureUsersSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var uname = String(username || "").trim();
  if (!uname) throw new Error("username is required.");

  var rowNum = findRowByUsername(sheet, uname);
  var idCol = headers.indexOf("id") + 1;
  var createdCol = headers.indexOf("created_at") + 1;
  var existingId = rowNum > 0 && idCol > 0 ? sheet.getRange(rowNum, idCol).getValue() : "";
  var existingCreated = rowNum > 0 && createdCol > 0 ? sheet.getRange(rowNum, createdCol).getValue() : "";

  var values = headers.map(function (h) {
    if (h === "id") return existingId || generateUUID();
    if (h === "username") return uname;
    if (h === "display_name") return displayName || uname;
    if (h === "password") return password || "";
    if (h === "role") return role || "user";
    if (h === "suspended") return false;
    if (h === "created_at") return existingCreated || new Date().toISOString();
    return "";
  });

  if (rowNum > 0) {
    sheet.getRange(rowNum, 1, rowNum, values.length).setValues([values]);
    Logger.log("Updated user: " + uname + " (role: " + (role || "user") + ")");
  } else {
    sheet.appendRow(values);
    Logger.log("Created user: " + uname + " (role: " + (role || "user") + ")");
  }
  return { ok: true, username: uname, role: role || "user" };
}

/**
 * RUN ONCE: Extensions → Apps Script → select setupDefaultAdminUser → Run ▶
 * Adds/updates the default BFP admin account on the users tab.
 */
function setupDefaultAdminUser() {
  var sheet = ensureUsersSheet();
  ensureColumnExists(sheet, "email");
  var result = ensureAdminUser(
    "mfcariaga@bfp",
    "admin123",
    "MFC Ariaga",
    "admin"
  );
  var rowNum = findRowByUsername(sheet, "mfcariaga@bfp");
  if (rowNum > 0) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var emailCol = headers.indexOf("email") + 1;
    if (emailCol > 0) {
      sheet.getRange(rowNum, emailCol).setValue("mfcariaga@bfp");
    }
  }
  Logger.log("✅ Default admin ready. Sign in with username: mfcariaga@bfp");
  return result;
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

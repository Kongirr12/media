/**
 * CODES.GS - Backend Logic (Object Data Support)
 * -------------------------------------------------------------------------
 * Update:
 * - [PERFORMANCE] Implemented Server-side Pagination & Searching
 * - [CONFIG] Page Limit setting mapped to Sheet 'Settings' Cell B3
 * - [V8] Optimized array operations for large datasets
 * - [FIX] getAllData now returns only the first page of resources
 * -------------------------------------------------------------------------
 */

// =========================================================================
// 1. CORE & CONFIGURATION
// =========================================================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('คลังสื่อการเรียนรู้ (Learning Resources)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAuthToken() {
  return ScriptApp.getOAuthToken();
}

function getSettings() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return {};
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return {};

    var data = sheet.getDataRange().getDisplayValues();
    var settings = {};
    
    // Read Key-Value pairs (Existing logic)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) settings[data[i][0]] = data[i][1];
    }
    
    // [NEW] Force read Page Limit from Cell B3
    // B3 value overrides any 'PageLimit' key found in the list
    var b3Value = sheet.getRange("B3").getValue();
    settings['PageLimit'] = (b3Value && !isNaN(b3Value)) ? parseInt(b3Value) : 20;

    return settings;
  } catch (e) {
    console.error("Error getting settings: " + e.toString());
    return { 'PageLimit': 20 };
  }
}

function _getDriveFolderId() {
  var settings = getSettings();
  if (!settings.DriveFolderId) {
    throw new Error("ไม่พบการตั้งค่า Drive Folder ID ในชีต Settings");
  }
  return settings.DriveFolderId;
}

// =========================================================================
// 2. DATABASE CONNECTION & HELPER
// =========================================================================

function getSheetData(sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 1) return [];

    var headers = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      var hasData = false;
      for (var j = 0; j < headers.length; j++) {
        var key = String(headers[j]).trim();
        if (key) {
          obj[key] = row[j];
          if (row[j]) hasData = true;
        }
      }
      if (hasData) result.push(obj);
    }
    return result;
  } catch (e) {
    console.log("Error getting sheet data: " + e.toString());
    return [];
  }
}

function getCombinedList(ss, listColumnName, legacySheetName) {
  var combined = [];
  var lookup = {};
  var listSheet = ss.getSheetByName('Lists');
  if (listSheet) {
    var listData = listSheet.getDataRange().getDisplayValues();
    if (listData.length > 0) {
      var headers = listData[0].map(function(h) { return String(h).trim().toLowerCase(); });
      var targetCol = String(listColumnName).toLowerCase();
      var colIdx = headers.indexOf(targetCol);
      if (colIdx > -1) {
        for (var i = 1; i < listData.length; i++) {
          var val = listData[i][colIdx];
          if (val && !lookup[val]) { combined.push(val); lookup[val] = true; }
        }
      }
    }
  }
  var legacyData = getSheetData(legacySheetName);
  for (var k = 0; k < legacyData.length; k++) {
    var row = legacyData[k];
    var val = row['list'] || row[Object.keys(row)[0]];
    if (val && !lookup[val]) { combined.push(val); lookup[val] = true; }
  }
  return combined;
}

// [NEW] Core Logic for Filtering and Pagination
function _filterAndPaginateResources(allResources, params, settings) {
  var page = parseInt(params.page) || 1;
  var limit = parseInt(params.limit) || parseInt(settings.PageLimit) || 20;
  var search = (params.search || "").toLowerCase();
  var subject = params.subject || "";
  var grade = params.grade || "";
  var type = params.type || "";
  var sort = params.sort || "newest";
  var favorites = params.favorites || []; // Array of favorited IDs
  var onlyFavorites = params.onlyFavorites === true || params.onlyFavorites === "true";

  // 1. Filter
  var filtered = allResources.filter(function(item) {
    var matchSearch = !search || (item.name || "").toLowerCase().includes(search) || (item.creator || "").toLowerCase().includes(search);
    var matchSubject = !subject || item.subject === subject;
    var matchGrade = !grade || (item.grade || "").includes(grade);
    var matchType = !type || item.type === type;
    var matchFav = !onlyFavorites || favorites.includes(String(item.id));

    return matchSearch && matchSubject && matchGrade && matchType && matchFav;
  });

  // 2. Sort
  filtered.sort(function(a, b) {
    var dateA = new Date(a.timestamp || 0);
    var dateB = new Date(b.timestamp || 0);
    
    if (sort === 'newest') return dateB - dateA;
    if (sort === 'oldest') return dateA - dateB;
    if (sort === 'name_asc') return (a.name || "").localeCompare(b.name || "", 'th');
    if (sort === 'name_desc') return (b.name || "").localeCompare(a.name || "", 'th');
    return 0;
  });

  // 3. Paginate
  var totalItems = filtered.length;
  var totalPages = Math.ceil(totalItems / limit);
  var startIndex = (page - 1) * limit;
  var endIndex = startIndex + limit;
  var paginatedData = filtered.slice(startIndex, endIndex);

  return {
    data: paginatedData,
    pagination: {
      page: page,
      limit: limit,
      totalItems: totalItems,
      totalPages: totalPages
    }
  };
}

// =========================================================================
// 3. API FOR FRONTEND
// =========================================================================

/**
 * คำนวณสถิติสำหรับ Dashboard
 * @param {Array} allResources - ข้อมูลสื่อทั้งหมด
 * @param {Array} subjects - รายการรายวิชา (ไม่ใช้แล้ว)
 * @param {Array} types - รายการประเภทสื่อ (ไม่ใช้แล้ว)
 * @returns {Object} stats object
 */
function _calculateStats(allResources, subjects, types) {
  // นับ Unique Creators, Subjects, Types จากข้อมูลสื่อจริง
  var creatorSet = {};
  var subjectSet = {};
  var typeSet = {};
  
  for (var i = 0; i < allResources.length; i++) {
    var resource = allResources[i];
    
    // นับ Unique Creators
    var creator = resource.creator;
    if (creator && creator.toString().trim() !== '') {
      creatorSet[creator] = true;
    }
    
    // นับ Unique Subjects
    var subject = resource.subject;
    if (subject && subject.toString().trim() !== '') {
      subjectSet[subject] = true;
    }
    
    // นับ Unique Types
    var type = resource.type;
    if (type && type.toString().trim() !== '') {
      typeSet[type] = true;
    }
  }

  return {
    totalResources: allResources.length,
    uniqueCreators: Object.keys(creatorSet).length,
    totalSubjects: Object.keys(subjectSet).length,
    totalTypes: Object.keys(typeSet).length
  };
}


function getAllData(initialParams) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settings = getSettings(); // Now includes PageLimit from B3
    
    // Fetch raw data
    var rawResources = getSheetData('Resources');
    var users = getSheetData('Users');
    var rawFavorites = getSheetData('Favorites');

    // Default params for first load
    var params = initialParams || { page: 1, limit: settings.PageLimit };
    
    // Perform initial pagination
    var resourceResult = _filterAndPaginateResources(rawResources, params, settings);

    // Get lists for dropdowns and stats
    var subjects = getCombinedList(ss, 'Subjects', 'Subjects');
    var grades = getCombinedList(ss, 'Grades', 'Grades');
    var types = getCombinedList(ss, 'Types', 'Types');
    
    // Calculate stats from ALL resources (not just paginated)
    var stats = _calculateStats(rawResources, subjects, types);

    return {
      resources: resourceResult.data, // Only Page 1
      pagination: resourceResult.pagination, // Meta data
      users: users,
      favorites: rawFavorites,
      subjects: subjects,
      grades: grades,
      types: types,
      settings: settings,
      stats: stats  // Dashboard statistics
    };
  } catch (e) {
    console.error("getAllData Error: " + e.toString());
    return { resources: [], users: [], favorites: [], subjects: [], grades: [], types: [], settings: {}, stats: {} };
  }
}


// [NEW] Dedicated Search API
function searchResources(params) {
  try {
    var rawResources = getSheetData('Resources');
    var settings = getSettings();
    return _filterAndPaginateResources(rawResources, params, settings);
  } catch(e) {
    return { data: [], pagination: { totalItems: 0, page: 1 } };
  }
}

function loginUser(username, password) {
  try {
    var users = getSheetData('Users');
    var user = null;
    for (var i = 0; i < users.length; i++) {
      if (String(users[i].username) === String(username) && String(users[i].password) === String(password)) {
        user = users[i];
        break;
      }
    }
    if (user) return { status: true, user: user };
    return { status: false, message: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) {
    return { status: false, message: "Server Error: " + e.toString() };
  }
}

// =========================================================================
// 4. RESOURCE MANAGEMENT
// =========================================================================

function _processUploadedFile(fileId, targetFolder, newNamePrefix) {
  try {
    var file = DriveApp.getFileById(fileId);
    if (newNamePrefix) {
       var ext = file.getName().split('.').pop();
       file.setName(newNamePrefix + "_" + file.getName());
    }
   // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (targetFolder) {
       file.moveTo(targetFolder);
    }
    return file.getId();
  } catch (e) {
    console.error("Process File Error (" + fileId + "): " + e.toString());
    return fileId;
  }
}

function saveResource(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Resources');
    var mainFolderId = _getDriveFolderId();
    var mainFolder = DriveApp.getFolderById(mainFolderId);

    var coverId = "";
    if (payload.coverFile && payload.coverFile.id) {
      coverId = _processUploadedFile(payload.coverFile.id, mainFolder, "Cover_" + payload.name);
    }

    var albumData = []; 
    var albumFolderId = "";
    
    if (payload.albumFiles && payload.albumFiles.length > 0) {
      var subFolder = mainFolder.createFolder("Album_" + payload.name + "_" + new Date().getTime());
      //subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      albumFolderId = subFolder.getId();

      for (var i = 0; i < payload.albumFiles.length; i++) {
        var f = payload.albumFiles[i];
        if (f.id) {
           var processedId = _processUploadedFile(f.id, subFolder, null);
           albumData.push({
             id: processedId,
             name: f.name || 'Unknown',
             type: f.type || 'application/octet-stream'
           });
        }
      }
    }

    var id = new Date().getTime().toString();
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var gradesStr = Array.isArray(payload.grades) ? payload.grades.join(", ") : (payload.grades || "");
    var albumJson = JSON.stringify(albumData);

    sheet.appendRow([
      id,
      payload.name,
      payload.subject,
      gradesStr,
      payload.type,
      payload.creator || 'Unknown',
      payload.url,
      coverId,
      albumFolderId,
      albumJson,
      payload.description,
      timestamp,
      0 
    ]);

    if (sheet.getLastRow() === 2) { 
       var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
       if (headers.length < 13) sheet.getRange(1, 13).setValue("Views");
    }

    // Return paginated data context (Page 1) to update UI immediately
    return { status: true, message: "บันทึกข้อมูลสำเร็จ", data: getAllData() };

  } catch (e) {
    return { status: false, message: "Error: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function editResource(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Resources');
    var data = sheet.getDataRange().getDisplayValues();
    var mainFolderId = _getDriveFolderId();
    var mainFolder = DriveApp.getFolderById(mainFolderId);

    var rowIndex = -1;
    var rowData = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        rowIndex = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1) return { status: false, message: "ไม่พบข้อมูล" };

    var coverId = rowData[7];
    if (payload.coverFile && payload.coverFile.id) {
      coverId = _processUploadedFile(payload.coverFile.id, mainFolder, "Cover_" + payload.name);
    } else if (payload.deleteCover === true) {
      coverId = "";
    }

    var albumFolderId = rowData[8];
    var currentAlbumData = [];
    try {
      currentAlbumData = JSON.parse(rowData[9] || "[]");
      if (!Array.isArray(currentAlbumData)) currentAlbumData = [];
    } catch(e) { currentAlbumData = []; }

    if (payload.deletedAlbumFiles && payload.deletedAlbumFiles.length > 0) {
      payload.deletedAlbumFiles.forEach(function(delId) {
        var idx = -1;
        for (var k = 0; k < currentAlbumData.length; k++) {
           var item = currentAlbumData[k];
           var itemId = (typeof item === 'object' && item !== null) ? item.id : item;
           if (String(itemId) === String(delId)) { idx = k; break; }
        }
        if (idx > -1) {
          currentAlbumData.splice(idx, 1);
          try { DriveApp.getFileById(delId).setTrashed(true); } catch(e) {}
        }
      });
    }

    if (payload.newAlbumFiles && payload.newAlbumFiles.length > 0) {
      var subFolder;
      if (albumFolderId) { try { subFolder = DriveApp.getFolderById(albumFolderId); } catch(e) {} }
      if (!subFolder) {
        subFolder = mainFolder.createFolder("Album_" + payload.name + "_" + new Date().getTime());
        //subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        albumFolderId = subFolder.getId();
      }
      for (var j = 0; j < payload.newAlbumFiles.length; j++) {
        var nf = payload.newAlbumFiles[j];
        if (nf.id) {
           var processedId = _processUploadedFile(nf.id, subFolder, null);
           currentAlbumData.push({ id: processedId, name: nf.name || 'Unknown', type: nf.type || 'application/octet-stream' });
        }
      }
    }

    var gradesStr = Array.isArray(payload.grades) ? payload.grades.join(", ") : (payload.grades || "");
    var albumJson = JSON.stringify(currentAlbumData);
    var currentViews = rowData[12] || 0;

    var newRow = [
      rowData[0],
      payload.name,
      payload.subject,
      gradesStr,
      payload.type,
      rowData[5],
      payload.url,
      coverId,
      albumFolderId,
      albumJson,
      payload.description,
      rowData[11],
      currentViews 
    ];

    sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
    
    // Return to first page to see changes (or could be enhanced to stay on page)
    return { status: true, message: "อัปเดตข้อมูลสำเร็จ", data: getAllData() };

  } catch (e) {
    return { status: false, message: "Error: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function incrementView(id) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(2000)) { 
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resources');
      var data = sheet.getDataRange().getDisplayValues();
      var headers = data[0];
      var viewColIdx = headers.indexOf('Views');

      if (viewColIdx === -1) {
        viewColIdx = headers.length;
        sheet.getRange(1, viewColIdx + 1).setValue('Views');
      }

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
           var cell = sheet.getRange(i + 1, viewColIdx + 1);
           var val = parseInt(cell.getValue());
           if (isNaN(val)) val = 0;
           cell.setValue(val + 1);
           return; 
        }
      }
    } catch(e) {
      console.error("Inc View Error: " + e);
    } finally {
      lock.releaseLock();
    }
  }
}

function toggleFavorite(userId, resourceId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Favorites');
    
    if (!sheet) {
      sheet = ss.insertSheet('Favorites');
      sheet.appendRow(['UserId', 'ResourceId', 'Timestamp']);
    }

    var data = sheet.getDataRange().getDisplayValues();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId) && String(data[i][1]) === String(resourceId)) {
        rowIndex = i + 1;
        break;
      }
    }

    var isFavorited = false;
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex); 
      isFavorited = false;
    } else {
      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.appendRow([userId, resourceId, timestamp]);
      isFavorited = true;
    }

    // Only return favorite list, not all resources to keep it light
    return { status: true, isFavorited: isFavorited, favorites: getSheetData('Favorites') };

  } catch (e) {
    return { status: false, message: "Error: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteResource(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Resources');
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { status: true, data: getAllData() };
      }
    }
    return { status: false, message: "ไม่พบข้อมูล" };
  } catch (e) {
    return { status: false, message: "Error" };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// 5. ADMIN UTILITIES
// =========================================================================

function updateMasterData(category, action, value, newValue) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Lists');
    if (!sheet) { sheet = ss.insertSheet('Lists'); sheet.appendRow(['Subjects', 'Grades', 'Types']); }
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0] || [];
    var colIndex = headers.map(function(h) { return String(h).trim().toLowerCase(); }).indexOf(String(category).toLowerCase());

    if (colIndex === -1) {
      if (action === 'add') {
        colIndex = headers.length;
        sheet.getRange(1, colIndex + 1).setValue(category);
      } else return getAllData();
    }

    var colValues = [];
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var range = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
      var vals = range.getDisplayValues();
      for (var k = 0; k < vals.length; k++) { if (vals[k][0] !== "") colValues.push(vals[k][0]); }
    }

    if (action === 'add') { if (colValues.indexOf(value) === -1) colValues.push(value); }
    else if (action === 'delete') { var idx = colValues.indexOf(value); if (idx > -1) colValues.splice(idx, 1); }
    else if (action === 'edit') { var idx = colValues.indexOf(value); if (idx > -1) colValues[idx] = newValue; }

    if (sheet.getMaxRows() > 1) sheet.getRange(2, colIndex + 1, sheet.getMaxRows() - 1, 1).clearContent();
    if (colValues.length > 0) {
      var output = colValues.map(function(v) { return [v]; });
      sheet.getRange(2, colIndex + 1, output.length, 1).setValues(output);
    }
    return getAllData();
  } catch (e) { throw new Error("Update Master Failed"); } finally { lock.releaseLock(); }
}

function addUser(userData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) { if (data[i][1] == userData.username) return { status: false, message: "Username ซ้ำ" }; }
    var newId = "U-" + new Date().getTime();
    sheet.appendRow([newId, userData.username, userData.password, userData.name, userData.role, userData.image]);
    return { status: true, data: getAllData() };
  } catch (e) { return { status: false, message: "Error: " + e.toString() }; } finally { lock.releaseLock(); }
}

function editUser(previousUsername, userData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data = sheet.getDataRange().getDisplayValues();
    var rowIndex = -1;
    if (previousUsername !== userData.username) {
      for (var i = 1; i < data.length; i++) { if (data[i][1] == userData.username) return { status: false, message: "Username ใหม่ซ้ำ" }; }
    }
    for (var i = 1; i < data.length; i++) { if (data[i][1] == previousUsername) { rowIndex = i + 1; break; } }
    if (rowIndex === -1) return { status: false, message: "ไม่พบผู้ใช้งาน" };
    sheet.getRange(rowIndex, 2, 1, 5).setValues([[userData.username, userData.password, userData.name, userData.role, userData.image]]);
    return { status: true, data: getAllData() };
  } catch (e) { return { status: false, message: "Error" }; } finally { lock.releaseLock(); }
}

function deleteUser(username) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == username) { sheet.deleteRow(i + 1); return { status: true, data: getAllData() }; }
    }
    return { status: false, message: "ไม่พบผู้ใช้งาน" };
  } catch (e) { return { status: false, message: "Error" }; } finally { lock.releaseLock(); }
}

function saveGeneralSettings(formObject) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { status: false, message: "ไม่พบ Sheet Settings" };
    
    // 1. Save normal settings (Key-Value)
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      if (formObject[key] !== undefined) { sheet.getRange(i + 1, 2).setValue(formObject[key]); }
    }

    // 2. [NEW] Save Page Limit to Cell B3 specifically
    if (formObject.PageLimit) {
      sheet.getRange("B3").setValue(formObject.PageLimit);
    }

    return { status: true, message: "บันทึกการตั้งค่าเรียบร้อย", data: getAllData() };
  } catch (e) { return { status: false, message: "Error" }; } finally { lock.releaseLock(); }
}
function forceDriveAuth() {
  // ฟังก์ชันหลอกให้ระบบรู้ว่าเราต้องการสิทธิ์ DriveApp
  var temp = DriveApp.getRootFolder();
}
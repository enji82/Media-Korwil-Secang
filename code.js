// File: Code.gs

/**
 * ===================================================================
 * ===================== 2. FUNGSI INTI APLIKASI =====================
 * ===================================================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * ===================================================================
 * ===================== 3. FUNGSI UTILITAS UMUM =====================
 * ===================================================================
 */

function handleError(functionName, error) {
  Logger.log(`Error di ${functionName}: ${error.message}\nStack: ${error.stack}`);
  throw new Error(error.message);
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

// ================== FUNGSI YANG HILANG ADA DI SINI ==================
function getDataFromSheet(configKey) {
  try {
    const config = SPREADSHEET_CONFIG[configKey];
    if (!config) throw new Error(`Konfigurasi untuk '${configKey}' tidak ditemukan.`);
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error(`Sheet '${config.sheet}' di spreadsheet '${config.id}' tidak ditemukan.`);
    return sheet.getDataRange().getValues();
  } catch (e) {
    handleError(`getDataFromSheet: ${configKey}`, e);
  }
}
// =====================================================================

function getCachedData(key, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached != null) {
    Logger.log("Mengambil data dari Cache: " + key);
    return JSON.parse(cached);
  }
  Logger.log("Mengambil data dari Spreadsheet dan menyimpan ke Cache: " + key);
  const freshData = fetchFunction();
  cache.put(key, JSON.stringify(freshData), 21600); // Simpan selama 6 jam
  return freshData;
}

function getFolders(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const subFolders = parentFolder.getFolders();
    const folderList = [];
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      folderList.push({
        id: folder.getId(),
        name: folder.getName()
      });
    }
    folderList.sort((a, b) => b.name.localeCompare(a.name));
    return folderList;
  } catch (e) {
    return handleError("getFolders", e);
  }
}

function getFiles(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const files = parentFolder.getFiles();
    const fileList = [];
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        url: file.getUrl()
      });
    }
    fileList.sort((a, b) => a.name.localeCompare(b.name));
    return fileList;
  } catch (e) {
    return handleError("getFiles", e);
  }
}
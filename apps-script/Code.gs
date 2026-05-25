const SPREADSHEET_ID = '1AFff64g1PR6qtcy9YzNR6xl1GNRbXFg33t8XPm3YyJw';
const FORM_DATA_SHEET_NAME = 'formData';
const CONFIG_SHEET_NAME = 'Config';

const DEFAULT_CONFIG = {
  isRegistrationOpen: true,
  enableQuestionnaire: true,
  eventName: 'Mellow.Sundae Weekly Activity',
  allowFollowers: true
};

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';

  if (action === 'get_config') {
    return outputConfig(e);
  }

  return HtmlService.createHtmlOutputFromFile('form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function submitForm(data) {
  const sheet = getOrCreateSheet(FORM_DATA_SHEET_NAME, []);

  sheet.appendRow([
    new Date(),
    data.activity,
    data.fullname,
    data.phone,
    data.followers,
    data.ig,
    data.first_time,
    data.misc
  ]);

  return '✅ ขอบคุณที่ลงทะเบียน!';
}

function doPost(e) {
  const action = e && e.parameter ? e.parameter.action : '';

  if (action === 'save_config') {
    try {
      return saveConfig(e);
    } catch (err) {
      return ContentService.createTextOutput('❌ Error: ' + err.message);
    }
  }

  return saveRegistration(e);
}

function saveRegistration(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getOrCreateSheet(FORM_DATA_SHEET_NAME, []);
    const data = JSON.parse(JSON.stringify(e.parameter));
    data.Timestamp = new Date();

    const lastCol = sheet.getLastColumn();
    let headers = [];

    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(Boolean);
    }

    Object.keys(data).forEach(function(key) {
      if (headers.indexOf(key) === -1) {
        headers.push(key);
        sheet.getRange(1, headers.length).setValue(key);
      }
    });

    const newRow = headers.map(function(header) {
      return data[header] || '';
    });

    sheet.appendRow(newRow);

    return ContentService.createTextOutput('✅ ขอบคุณที่ลงทะเบียน!');
  } catch (err) {
    return ContentService.createTextOutput('❌ Error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

function outputConfig(e) {
  const config = getConfig();
  const callback = e.parameter.callback || '';

  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback');
    }

    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(config) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return outputJson(config);
}

function outputJson(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getConfig() {
  const sheet = getOrCreateSheet(CONFIG_SHEET_NAME, ['key', 'value']);
  const values = sheet.getDataRange().getValues();
  const config = Object.assign({}, DEFAULT_CONFIG);

  for (let row = 1; row < values.length; row++) {
    const key = values[row][0];
    const value = values[row][1];
    if (!key) continue;
    config[key] = parseConfigValue(value);
  }

  ensureDefaultConfigRows(sheet, config);
  return config;
}

function parseConfigValue(value) {
  if (value === true || value === 'true') return true;
  if (value === false || value === 'false') return false;
  return value;
}

function saveConfig(e) {
  assertAdminToken(e.parameter.adminToken || '');

  const config = JSON.parse(e.parameter.config || '{}');
  const safeConfig = {
    isRegistrationOpen: config.isRegistrationOpen === true,
    enableQuestionnaire: config.enableQuestionnaire === true,
    eventName: String(config.eventName || '').trim(),
    allowFollowers: config.allowFollowers === true
  };

  if (!safeConfig.eventName) {
    throw new Error('eventName is required');
  }

  const sheet = getOrCreateSheet(CONFIG_SHEET_NAME, ['key', 'value']);
  sheet.clearContents();
  sheet.appendRow(['key', 'value']);

  Object.keys(safeConfig).forEach(function(key) {
    sheet.appendRow([key, safeConfig[key]]);
  });

  return outputJson({ ok: true, config: safeConfig });
}

function assertAdminToken(token) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');

  if (!expectedToken) {
    throw new Error('ADMIN_TOKEN script property is not set');
  }

  if (token !== expectedToken) {
    throw new Error('Invalid admin token');
  }
}

function ensureDefaultConfigRows(sheet, config) {
  if (sheet.getLastRow() > 1) return;

  sheet.clearContents();
  sheet.appendRow(['key', 'value']);

  Object.keys(config).forEach(function(key) {
    sheet.appendRow([key, config[key]]);
  });
}

function getOrCreateSheet(name, headers) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  if (headers.length && sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  return sheet;
}

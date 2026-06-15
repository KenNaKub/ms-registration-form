const SPREADSHEET_ID = '1AFff64g1PR6qtcy9YzNR6xl1GNRbXFg33t8XPm3YyJw';
const FORM_DATA_SHEET_NAME = 'formData';
const CONFIG_SHEET_NAME = 'Config';
const MAIN_QUESTIONS_SHEET_NAME = 'MainQuestions';
const QUESTIONNAIRE_QUESTIONS_SHEET_NAME = 'QuestionnaireQuestions';
const QUESTION_HEADERS = ['order', 'name', 'label', 'type', 'required', 'readonly', 'placeholder', 'options'];

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

  if (action === 'get_form_definition') {
    return outputFormDefinition(e);
  }

  if (action === 'save_form_definition') {
    return outputSaveFormDefinition(e);
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

  if (action === 'save_form_definition') {
    try {
      return outputJson(saveFormDefinitionData(e));
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

  return outputJsonp(config, callback);
}

function outputFormDefinition(e) {
  const definition = getFormDefinition();
  const callback = e.parameter.callback || '';

  return outputJsonp(definition, callback);
}

function outputSaveFormDefinition(e) {
  const callback = e.parameter.callback || '';

  try {
    return outputJsonp(saveFormDefinitionData(e), callback);
  } catch (err) {
    return outputJsonp({ ok: false, error: err.message }, callback);
  }
}

function outputJsonp(payload, callback) {
  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback');
    }

    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return outputJson(payload);
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

function getFormDefinition() {
  const config = getConfig();

  return {
    config: config,
    main: {
      title: config.mainTitle || '',
      subtitle: config.mainSubtitle || '',
      questions: getQuestionsFromSheet(MAIN_QUESTIONS_SHEET_NAME)
    },
    questionnaire: {
      title: config.questionnaireTitle || '',
      questions: getQuestionsFromSheet(QUESTIONNAIRE_QUESTIONS_SHEET_NAME)
    }
  };
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

function saveFormDefinitionData(e) {
  assertAdminToken(e.parameter.adminToken || '');

  const definition = JSON.parse(e.parameter.definition || '{}');
  const config = definition.config || {};
  const main = definition.main || {};
  const questionnaire = definition.questionnaire || {};

  const safeConfig = {
    isRegistrationOpen: config.isRegistrationOpen === true,
    enableQuestionnaire: config.enableQuestionnaire === true,
    eventName: String(config.eventName || '').trim(),
    allowFollowers: config.allowFollowers === true,
    mainTitle: String(main.title || '').trim(),
    mainSubtitle: String(main.subtitle || '').trim(),
    questionnaireTitle: String(questionnaire.title || '').trim()
  };

  if (!safeConfig.eventName) {
    throw new Error('eventName is required');
  }

  const mainQuestions = sanitizeQuestions(main.questions || []);
  const questionnaireQuestions = sanitizeQuestions(questionnaire.questions || []);

  saveConfigValues(safeConfig);
  saveQuestionsToSheet(MAIN_QUESTIONS_SHEET_NAME, mainQuestions);
  saveQuestionsToSheet(QUESTIONNAIRE_QUESTIONS_SHEET_NAME, questionnaireQuestions);
  SpreadsheetApp.flush();

  return {
    ok: true,
    definition: {
      config: safeConfig,
      main: {
        title: safeConfig.mainTitle,
        subtitle: safeConfig.mainSubtitle,
        questions: mainQuestions
      },
      questionnaire: {
        title: safeConfig.questionnaireTitle,
        questions: questionnaireQuestions
      }
    }
  };
}

function saveConfigValues(config) {
  const sheet = getOrCreateSheet(CONFIG_SHEET_NAME, ['key', 'value']);
  sheet.clearContents();
  sheet.appendRow(['key', 'value']);

  Object.keys(config).forEach(function(key) {
    sheet.appendRow([key, config[key]]);
  });
}

function sanitizeQuestions(questions) {
  return questions.map(function(question) {
    const type = String(question.type || 'text').trim();
    const safeQuestion = {
      name: String(question.name || '').trim(),
      label: String(question.label || '').trim(),
      type: type
    };

    if (!/^[A-Za-z][0-9A-Za-z_]*$/.test(safeQuestion.name)) {
      throw new Error('Invalid question name: ' + safeQuestion.name);
    }

    if (!safeQuestion.label) {
      throw new Error('Question label is required for: ' + safeQuestion.name);
    }

    if (question.required === true) safeQuestion.required = true;
    if (question.readonly === true) safeQuestion.readonly = true;

    if (question.placeholder) {
      safeQuestion.placeholder = String(question.placeholder).trim();
    }

    if (['select', 'radio', 'checkbox'].indexOf(type) !== -1) {
      safeQuestion.options = sanitizeOptions(question.options || []);
      if (!safeQuestion.options.length) {
        throw new Error('Options are required for: ' + safeQuestion.name);
      }
    }

    return safeQuestion;
  });
}

function sanitizeOptions(options) {
  return options.map(function(option) {
    if (typeof option === 'string') return option.trim();

    return {
      value: String(option.value || '').trim(),
      label: String(option.label || '').trim()
    };
  }).filter(function(option) {
    if (typeof option === 'string') return option;
    return option.value && option.label;
  });
}

function getQuestionsFromSheet(sheetName) {
  const sheet = getOrCreateSheet(sheetName, QUESTION_HEADERS);

  if (sheet.getLastRow() <= 1) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, QUESTION_HEADERS.length).getValues();

  return values.map(function(row) {
    const question = {
      name: row[1],
      label: row[2],
      type: row[3] || 'text'
    };

    if (parseConfigValue(row[4]) === true) question.required = true;
    if (parseConfigValue(row[5]) === true) question.readonly = true;
    if (row[6]) question.placeholder = row[6];

    if (row[7]) {
      question.options = JSON.parse(row[7]);
    }

    return question;
  }).filter(function(question) {
    return question.name && question.label;
  });
}

function saveQuestionsToSheet(sheetName, questions) {
  const sheet = getOrCreateSheet(sheetName, QUESTION_HEADERS);
  sheet.clearContents();
  sheet.appendRow(QUESTION_HEADERS);

  questions.forEach(function(question, index) {
    sheet.appendRow([
      index + 1,
      question.name,
      question.label,
      question.type,
      question.required === true,
      question.readonly === true,
      question.placeholder || '',
      question.options ? JSON.stringify(question.options) : ''
    ]);
  });
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

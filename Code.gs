/**
 * RMA Management for Google Sheets
 *
 * Workflow:
 * NewEntry sheet manual input -> Jitbit menu -> issue Entry ID
 * -> append ledger -> copy and populate RMA Registration Form template.
 */

const CONFIG = {
  spreadsheetId: '1uGWkEmfg0B00mGANWSVLLMgMUWKEiKIbHncDZhRzBzs',
  templateDocumentId: '1CSSxjyQbFN4HYw4uHJnPS9NqOTRvIsc9',
  initSheetName: '__init',
  newEntrySheetName: 'NewEntry',
  ledgerSheetName: 'RMA Ledger',
  outputFolderName: 'RMA Generated Forms',
  menuName: '\uD83D\uDC8EJitbit',
  idPrefix: 'RMA',
  maxGoodsRows: 5,
  clearAfterSubmit: false,
};

const ENTRY_FIELDS = [
  { key: 'company', label: 'Company or institution', required: true },
  { key: 'contactPerson', label: 'Contact person', required: true },
  { key: 'street', label: 'Street' },
  { key: 'postalCode', label: 'Postal code' },
  { key: 'municipality', label: 'Municipality' },
  { key: 'country', label: 'Country' },
  { key: 'tel', label: 'Tel.' },
  { key: 'email', label: 'Email', required: true },
  { key: 'returnCompany', label: 'Return Company or institution' },
  { key: 'returnContactPerson', label: 'Return Contact person' },
  { key: 'returnStreet', label: 'Return Street' },
  { key: 'returnPostalCode', label: 'Return Postal code' },
  { key: 'returnMunicipality', label: 'Return Municipality' },
  { key: 'returnCountry', label: 'Return Country' },
  { key: 'returnTel', label: 'Return Tel.' },
  { key: 'returnEmail', label: 'Return Email' },
  { key: 'televicDeliveryNote', label: 'Televic Delivery Note' },
  { key: 'televicInvoiceNumber', label: 'Televic Invoice Number' },
  { key: 'installationProject', label: 'Installation / Project' },
  { key: 'reason', label: 'Reason' },
  { key: 'additionalRemarks', label: 'Additional remarks' },
];

const GOODS_FIELDS = [
  { key: 'tlvCode', label: 'TLV Code' },
  { key: 'partNumber', label: 'Part Number' },
  { key: 'description', label: 'Description' },
  { key: 'serialNumber', label: 'Serial Number (batch No. + Serial No.)' },
  { key: 'quantity', label: 'Quantity' },
  { key: 'errorDescription', label: 'Error description' },
];

const LEDGER_HEADERS = [
  'Entry ID',
  'Created at',
  'Company or institution',
  'Contact person',
  'Email',
  'Tel.',
  'Installation / Project',
  'Reason',
  'First item',
  'Form URL',
  'Document ID',
  'Status',
  'Notes',
];

const PLACEHOLDER_MAP = {
  entryId: ['{{ENTRY_ID}}', '{{YOUR_REF}}', '{{Your ref.}}'],
  dateOfApplication: ['{{DATE}}', '{{DATE_OF_APPLICATION}}'],
  company: ['{{COMPANY}}'],
  contactPerson: ['{{CONTACT_PERSON}}'],
  street: ['{{STREET}}'],
  postalCode: ['{{POSTAL_CODE}}'],
  municipality: ['{{MUNICIPALITY}}'],
  country: ['{{COUNTRY}}'],
  tel: ['{{TEL}}'],
  email: ['{{EMAIL}}'],
  returnCompany: ['{{RETURN_COMPANY}}'],
  returnContactPerson: ['{{RETURN_CONTACT_PERSON}}'],
  returnStreet: ['{{RETURN_STREET}}'],
  returnPostalCode: ['{{RETURN_POSTAL_CODE}}'],
  returnMunicipality: ['{{RETURN_MUNICIPALITY}}'],
  returnCountry: ['{{RETURN_COUNTRY}}'],
  returnTel: ['{{RETURN_TEL}}'],
  returnEmail: ['{{RETURN_EMAIL}}'],
  televicDeliveryNote: ['{{TELEVIC_DELIVERY_NOTE}}'],
  televicInvoiceNumber: ['{{TELEVIC_INVOICE_NUMBER}}'],
  installationProject: ['{{INSTALLATION_PROJECT}}'],
  additionalRemarks: ['{{ADDITIONAL_REMARKS}}'],
  reason: ['{{REASON}}'],
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.menuName)
    .addItem('\u30D5\u30A9\u30FC\u30E0\u751F\u6210', 'generateRmaForm')
    .addSeparator()
    .addItem('\u521D\u671F\u30BB\u30C3\u30C8\u30A2\u30C3\u30D7', 'setupRmaSheets')
    .addToUi();
}

function setupRmaSheets() {
  const ss = getSpreadsheet_();
  const initSheet = getOrCreateSheet_(ss, CONFIG.initSheetName);
  const newEntrySheet = getOrCreateSheet_(ss, CONFIG.newEntrySheetName);
  const ledgerSheet = getOrCreateSheet_(ss, CONFIG.ledgerSheetName);

  setupInitSheet_(initSheet);
  setupNewEntrySheet_(newEntrySheet);
  setupLedgerSheet_(ledgerSheet);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('RMA sheet setup is complete.');
}

function generateRmaForm() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = getSpreadsheet_();
    const initSheet = ss.getSheetByName(CONFIG.initSheetName);
    const newEntrySheet = ss.getSheetByName(CONFIG.newEntrySheetName);
    const ledgerSheet = ss.getSheetByName(CONFIG.ledgerSheetName);

    if (!initSheet || !newEntrySheet || !ledgerSheet) {
      throw new Error('__init, NewEntry, or RMA Ledger sheet was not found. Run setup first.');
    }

    const entry = readEntry_(initSheet, newEntrySheet);
    validateEntry_(entry);

    const entryId = issueEntryId_(ledgerSheet);
    const createdAt = new Date();
    const formFile = createRmaFormDocument_(entry, entryId, createdAt);

    appendLedgerRow_(ledgerSheet, entry, entryId, createdAt, formFile);

    if (CONFIG.clearAfterSubmit) {
      clearInputValues_(newEntrySheet);
    }

    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert(
      'RMA form was generated.\n\n' +
      'Entry ID: ' + entryId + '\n' +
      'Form URL: ' + formFile.getUrl()
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Failed to generate RMA form.\n\n' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function readEntry_(initSheet, goodsSheet) {
  return {
    fields: readInitFields_(initSheet),
    goods: readGoodsRows_(goodsSheet),
  };
}

function readInitFields_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const byLabel = {};

  values.slice(1).forEach(function(row) {
    const label = String(row[0] || '').trim();
    if (label) byLabel[label] = String(row[1] || '').trim();
  });

  const fields = {};
  ENTRY_FIELDS.forEach(function(field) {
    fields[field.key] = byLabel[field.label] || '';
  });

  return fields;
}

function readGoodsRows_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const headerRow = values[0] || [];
  const columnByLabel = {};
  const goods = [];

  headerRow.forEach(function(header, index) {
    if (header) columnByLabel[String(header).trim()] = index;
  });

  values.slice(1, CONFIG.maxGoodsRows + 1).forEach(function(row) {
    const item = {};
    GOODS_FIELDS.forEach(function(field) {
      const index = columnByLabel[field.label];
      item[field.key] = index === undefined ? '' : String(row[index] || '').trim();
    });
    if (hasAnyValue_(item)) goods.push(item);
  });

  return goods;
}

function validateEntry_(entry) {
  const missing = ENTRY_FIELDS
    .filter(function(field) {
      return field.required && !entry.fields[field.key];
    })
    .map(function(field) {
      return field.label;
    });

  if (missing.length > 0) {
    throw new Error('Required fields are missing: ' + missing.join(', '));
  }

  if (entry.goods.length === 0) {
    throw new Error('Enter at least one returned goods row.');
  }
}

function issueEntryId_(ledgerSheet) {
  const now = new Date();
  const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const prefix = CONFIG.idPrefix + '-' + datePart + '-';
  const lastRow = ledgerSheet.getLastRow();
  let maxSerial = 0;

  if (lastRow >= 2) {
    const ids = ledgerSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    ids.forEach(function(row) {
      const id = row[0] || '';
      if (id.indexOf(prefix) === 0) {
        const serial = Number(id.slice(prefix.length));
        if (!isNaN(serial)) maxSerial = Math.max(maxSerial, serial);
      }
    });
  }

  return prefix + String(maxSerial + 1).padStart(4, '0');
}

function createRmaFormDocument_(entry, entryId, createdAt) {
  const folder = getOrCreateOutputFolder_();
  const title = entryId + ' RMA Registration Form - ' + entry.fields.company;
  const templateFile = DriveApp.getFileById(CONFIG.templateDocumentId);
  const file = templateFile.makeCopy(title, folder);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  fillTemplate_(body, entry, entryId, createdAt);
  appendGeneratedData_(body, entry, entryId, createdAt);

  doc.saveAndClose();
  return file;
}

function fillTemplate_(body, entry, entryId, createdAt) {
  const values = buildTemplateValues_(entry, entryId, createdAt);

  Object.keys(PLACEHOLDER_MAP).forEach(function(key) {
    PLACEHOLDER_MAP[key].forEach(function(placeholder) {
      body.replaceText(escapeRegExp_(placeholder), values[key] || '');
    });
  });

  fillNextCellByLabel_(body, 'Date of the application', values.dateOfApplication);
  fillNextCellByLabel_(body, 'Your ref.', entryId);
  fillNextCellByLabel_(body, 'Company or institution', entry.fields.company);
  fillNextCellByLabel_(body, 'Contact person', entry.fields.contactPerson);
  fillNextCellByLabel_(body, 'Street', entry.fields.street);
  fillNextCellByLabel_(body, 'Postal code', entry.fields.postalCode);
  fillNextCellByLabel_(body, 'Municipality', entry.fields.municipality);
  fillNextCellByLabel_(body, 'Country', entry.fields.country);
  fillNextCellByLabel_(body, 'Tel.', entry.fields.tel);
  fillNextCellByLabel_(body, 'Email', entry.fields.email);
  fillNextCellByLabel_(body, 'Televic Delivery Note', entry.fields.televicDeliveryNote);
  fillNextCellByLabel_(body, 'Televic Invoice Number', entry.fields.televicInvoiceNumber);
  fillNextCellByLabel_(body, 'Installation / Project', entry.fields.installationProject);
  fillNextCellByLabel_(body, 'Additional remarks', entry.fields.additionalRemarks);
}

function appendGeneratedData_(body, entry, entryId, createdAt) {
  body.appendPageBreak();
  addTitle_(body, 'Generated RMA Entry Data');
  addKeyValueTable_(body, [
    ['Date of the application', formatDate_(createdAt)],
    ['Your ref.', entryId],
  ]);

  addSection_(body, 'Applicant data');
  addKeyValueTable_(body, [
    ['Company or institution', entry.fields.company],
    ['Contact person', entry.fields.contactPerson],
    ['Street', entry.fields.street],
    ['Postal code', entry.fields.postalCode],
    ['Municipality', entry.fields.municipality],
    ['Country', entry.fields.country],
    ['Tel.', entry.fields.tel],
    ['Email', entry.fields.email],
  ]);

  addSection_(body, 'Return address data');
  addKeyValueTable_(body, [
    ['Company or institution', entry.fields.returnCompany],
    ['Contact person', entry.fields.returnContactPerson],
    ['Street', entry.fields.returnStreet],
    ['Postal code', entry.fields.returnPostalCode],
    ['Municipality', entry.fields.returnMunicipality],
    ['Country', entry.fields.returnCountry],
    ['Tel.', entry.fields.returnTel],
    ['Email', entry.fields.returnEmail],
  ]);

  addSection_(body, 'Data concerning the goods that are to be returned');
  addGoodsTable_(body, entry.goods);

  addSection_(body, 'Other data');
  addKeyValueTable_(body, [
    ['Goods were originally delivered with Televic Delivery Note', entry.fields.televicDeliveryNote],
    ['Televic Invoice Number', entry.fields.televicInvoiceNumber],
    ['Installation / Project', entry.fields.installationProject],
    ['Additional remarks', entry.fields.additionalRemarks],
  ]);

  addSection_(body, 'Reasons for sending back the material');
  addReasons_(body, entry.fields.reason);
}

function buildTemplateValues_(entry, entryId, createdAt) {
  return {
    entryId: entryId,
    dateOfApplication: formatDate_(createdAt),
    company: entry.fields.company,
    contactPerson: entry.fields.contactPerson,
    street: entry.fields.street,
    postalCode: entry.fields.postalCode,
    municipality: entry.fields.municipality,
    country: entry.fields.country,
    tel: entry.fields.tel,
    email: entry.fields.email,
    returnCompany: entry.fields.returnCompany,
    returnContactPerson: entry.fields.returnContactPerson,
    returnStreet: entry.fields.returnStreet,
    returnPostalCode: entry.fields.returnPostalCode,
    returnMunicipality: entry.fields.returnMunicipality,
    returnCountry: entry.fields.returnCountry,
    returnTel: entry.fields.returnTel,
    returnEmail: entry.fields.returnEmail,
    televicDeliveryNote: entry.fields.televicDeliveryNote,
    televicInvoiceNumber: entry.fields.televicInvoiceNumber,
    installationProject: entry.fields.installationProject,
    additionalRemarks: entry.fields.additionalRemarks,
    reason: entry.fields.reason,
  };
}

function appendLedgerRow_(sheet, entry, entryId, createdAt, formFile) {
  setupLedgerSheet_(sheet);
  const firstItem = entry.goods[0] || {};

  sheet.appendRow([
    entryId,
    createdAt,
    entry.fields.company,
    entry.fields.contactPerson,
    entry.fields.email,
    entry.fields.tel,
    entry.fields.installationProject,
    entry.fields.reason,
    [firstItem.tlvCode, firstItem.partNumber, firstItem.description].filter(String).join(' / '),
    formFile.getUrl(),
    formFile.getId(),
    'New',
    '',
  ]);
}

function setupInitSheet_(sheet) {
  const existingValues = sheet.getDataRange().getDisplayValues();
  const valueByLabel = {};

  existingValues.slice(1).forEach(function(row) {
    const label = String(row[0] || '').trim();
    if (label) valueByLabel[label] = row[1] || '';
  });

  const rows = ENTRY_FIELDS.map(function(field) {
    return [field.label, valueByLabel[field.label] || ''];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['Field', 'Value']]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
  sheet.getRange(2, 2, rows.length, 1).setBackground('#fff8e1');

  const reasonRow = ENTRY_FIELDS.findIndex(function(field) {
    return field.key === 'reason';
  }) + 2;
  if (reasonRow > 1) {
    sheet.getRange(reasonRow, 2).setDataValidation(createReasonValidation_());
  }
}

function setupNewEntrySheet_(sheet) {
  const headers = ['Goods No'].concat(GOODS_FIELDS.map(function(field) {
    return field.label;
  }));
  const rows = [];

  for (let rowNumber = 1; rowNumber <= CONFIG.maxGoodsRows; rowNumber++) {
    rows.push(['Goods ' + rowNumber, '', '', '', '', '', '']);
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(2, 2, rows.length, headers.length - 1).setBackground('#fff8e1');
  sheet.getRange(2, 1, rows.length, 1).setFontWeight('bold').setBackground('#f3f6fb');
  sheet.getRange(2, 2).setNote('Enter returned goods here. Common entry data is managed in __init.');
}

function setupLedgerSheet_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).getDisplayValues()[0];
  const hasHeaders = firstRow.some(function(value) {
    return value;
  });

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]);
    sheet.setFrozenRows(1);
  }

  sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.autoResizeColumns(1, LEDGER_HEADERS.length);
}

function clearInputValues_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > 1) {
    sheet.getRange(2, 2, CONFIG.maxGoodsRows, lastColumn - 1).clearContent();
  }
}

function createReasonValidation_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'warranty repair',
      'repair out of warranty',
      'DOA (dead on arrival)',
      'Extended warranty applicable',
    ], true)
    .setAllowInvalid(true)
    .build();
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  return active || SpreadsheetApp.openById(CONFIG.spreadsheetId);
}

function getOrCreateSheet_(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function getOrCreateOutputFolder_() {
  const folders = DriveApp.getFoldersByName(CONFIG.outputFolderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.outputFolderName);
}

function hasAnyValue_(object) {
  return Object.keys(object).some(function(key) {
    return object[key] !== '';
  });
}

function addTitle_(body, text) {
  body.appendParagraph(text)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setBold(true);
}

function addSection_(body, text) {
  body.appendParagraph('');
  body.appendParagraph(text)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setBold(true);
}

function addKeyValueTable_(body, rows) {
  const table = body.appendTable(rows.map(function(row) {
    return [row[0], row[1] || ''];
  }));
  table.setBorderWidth(0.5);

  for (let i = 0; i < table.getNumRows(); i++) {
    const row = table.getRow(i);
    row.getCell(0).setBackgroundColor('#f3f6fb').setWidth(180);
    row.getCell(0).editAsText().setBold(true);
  }
}

function addGoodsTable_(body, goods) {
  const rows = [[
    'TLV Code',
    'Part Number',
    'Description',
    'Serial Number (batch No. + Serial No.)',
    'Quantity',
    'Error description',
  ]];

  goods.forEach(function(item) {
    rows.push([
      item.tlvCode || '',
      item.partNumber || '',
      item.description || '',
      item.serialNumber || '',
      item.quantity || '',
      item.errorDescription || '',
    ]);
  });

  while (rows.length <= CONFIG.maxGoodsRows) {
    rows.push(['', '', '', '', '', '']);
  }

  const table = body.appendTable(rows);
  table.setBorderWidth(0.5);

  const header = table.getRow(0);
  for (let i = 0; i < header.getNumCells(); i++) {
    header.getCell(i).setBackgroundColor('#f3f6fb');
    header.getCell(i).editAsText().setBold(true);
  }
}

function addReasons_(body, selectedReason) {
  const reasons = [
    'warranty repair',
    'repair out of warranty',
    'DOA (dead on arrival)',
    'Extended warranty applicable',
  ];

  reasons.forEach(function(reason) {
    const checked = normalize_(reason) === normalize_(selectedReason) ? '[x]' : '[ ]';
    body.appendParagraph(checked + ' ' + reason);
  });
}

function fillNextCellByLabel_(body, label, value) {
  if (!value) return false;
  const tables = body.getTables();
  for (let tableIndex = 0; tableIndex < tables.length; tableIndex++) {
    const table = tables[tableIndex];
    for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex++) {
      const row = table.getRow(rowIndex);
      for (let cellIndex = 0; cellIndex < row.getNumCells() - 1; cellIndex++) {
        const cell = row.getCell(cellIndex);
        if (normalizeLabel_(cell.getText()).indexOf(normalizeLabel_(label)) !== -1) {
          row.getCell(cellIndex + 1).setText(value);
          return true;
        }
      }
    }
  }
  return false;
}

function escapeRegExp_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeLabel_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function normalize_(value) {
  return String(value || '').trim().toLowerCase();
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

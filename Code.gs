/**
 * RMA Management for Google Sheets
 *
 * Workflow:
 * NewEntry sheet manual input -> 💎Jitbit > フォーム生成
 * -> entry ID issuance -> 管理台帳 append -> RMA form Google Doc generation.
 */

const CONFIG = {
  newEntrySheetName: 'NewEntry',
  ledgerSheetName: '管理台帳',
  outputFolderName: 'RMA Generated Forms',
  menuName: '💎Jitbit',
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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.menuName)
    .addItem('フォーム生成', 'generateRmaForm')
    .addSeparator()
    .addItem('初期セットアップ', 'setupRmaSheets')
    .addToUi();
}

function setupRmaSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newEntrySheet = getOrCreateSheet_(ss, CONFIG.newEntrySheetName);
  const ledgerSheet = getOrCreateSheet_(ss, CONFIG.ledgerSheetName);

  setupNewEntrySheet_(newEntrySheet);
  setupLedgerSheet_(ledgerSheet);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('RMA 管理シートの初期セットアップが完了しました。');
}

function generateRmaForm() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const newEntrySheet = ss.getSheetByName(CONFIG.newEntrySheetName);
    const ledgerSheet = ss.getSheetByName(CONFIG.ledgerSheetName);

    if (!newEntrySheet || !ledgerSheet) {
      throw new Error('NewEntry または 管理台帳 シートが見つかりません。先に「初期セットアップ」を実行してください。');
    }

    const entry = readNewEntry_(newEntrySheet);
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
      'フォーム生成が完了しました。\n\n' +
      'Entry ID: ' + entryId + '\n' +
      'Form URL: ' + formFile.getUrl()
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('フォーム生成に失敗しました。\n\n' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function readNewEntry_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const headerRow = values[0] || [];
  const inputRow = values[1] || [];
  const byLabel = {};

  headerRow.forEach(function(header, index) {
    if (header) byLabel[String(header).trim()] = String(inputRow[index] || '').trim();
  });

  const entry = {
    fields: {},
    goods: [],
  };

  ENTRY_FIELDS.forEach(function(field) {
    entry.fields[field.key] = byLabel[field.label] || '';
  });

  for (let rowNumber = 1; rowNumber <= CONFIG.maxGoodsRows; rowNumber++) {
    const item = {};
    GOODS_FIELDS.forEach(function(field) {
      item[field.key] = byLabel['Goods ' + rowNumber + ' - ' + field.label] || '';
    });
    if (hasAnyValue_(item)) entry.goods.push(item);
  }

  return entry;
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
    throw new Error('必須項目が未入力です: ' + missing.join(', '));
  }

  if (entry.goods.length === 0) {
    throw new Error('返送品情報を 1 行以上入力してください。');
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
  const doc = DocumentApp.create(title);
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  body.clear();

  addTitle_(body, 'RMA FORM (Return Material Authorization)');
  body.appendParagraph('RMA: To be completed by TELEVIC').setItalic(true);
  body.appendParagraph('TELEVIC Conference NV');
  body.appendParagraph('att. Customer Services');
  body.appendParagraph('Leo Bekaertlaan 1');
  body.appendParagraph('B-8870 IZEGEM (BELGIUM)');
  body.appendParagraph('');

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
  body.appendParagraph('(only to be filled out if different from applicant data)').setItalic(true);
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

  addKeyValueTable_(body, [
    ['Goods were originally delivered with Televic Delivery Note', entry.fields.televicDeliveryNote],
    ['Televic Invoice Number', entry.fields.televicInvoiceNumber],
    ['Installation / Project', entry.fields.installationProject],
    ['Additional remarks', entry.fields.additionalRemarks],
  ]);

  addSection_(body, 'Reasons for sending back the material');
  addReasons_(body, entry.fields.reason);

  doc.saveAndClose();
  return file;
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

function setupNewEntrySheet_(sheet) {
  const headers = ENTRY_FIELDS.map(function(field) {
    return field.label;
  });

  for (let rowNumber = 1; rowNumber <= CONFIG.maxGoodsRows; rowNumber++) {
    GOODS_FIELDS.forEach(function(field) {
      headers.push('Goods ' + rowNumber + ' - ' + field.label);
    });
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(2, 1, 1, headers.length).setBackground('#fff8e1');
  sheet.getRange(2, 1).setNote('ここに 1 件分を手入力してから 💎Jitbit > フォーム生成 を実行してください。');

  const reasonColumn = headers.indexOf('Reason') + 1;
  if (reasonColumn > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'warranty repair',
        'repair out of warranty',
        'DOA (dead on arrival)',
        'Extended warranty applicable',
      ], true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, reasonColumn).setDataValidation(rule);
  }
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
  if (lastColumn > 0) {
    sheet.getRange(2, 1, 1, lastColumn).clearContent();
  }
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
    const checked = normalize_(reason) === normalize_(selectedReason) ? '☑' : '☐';
    body.appendParagraph(checked + ' ' + reason);
  });
}

function normalize_(value) {
  return String(value || '').trim().toLowerCase();
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

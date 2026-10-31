/**
 * PERSONAL PEOPLE APP — Apps Script backend
 *
 * This script lives inside your Google Sheet (Extensions > Apps Script)
 * and is deployed as a Web App. The front-end (the artifact) calls these
 * functions over HTTP using fetch().
 *
 * SHEETS THIS SCRIPT EXPECTS (created automatically on first run if missing):
 *   - People            : id, name, categories, email, birthDate, imageUrl
 *   - Relationships      : id, personA, personB, type
 *   - About              : id, personId, date, text
 *   - FieldDefinitions   : id, label, type, important
 *   - FieldValues        : id, personId, fieldId, value
 *   - Categories         : name, imageUrl
 *   - Items              : id, collection, name, fields (JSON)
 *     (Collections: books / places / recipes. Item notes live in the About
 *      sheet with personId = item id; "Recommend to" links live in
 *      Relationships with personA = item id.)
 *
 * SECURITY: secrets are NOT stored in this file (the repo is public).
 * Set them in Apps Script: Project Settings (gear) > Script Properties:
 *   APP_KEY         = the password you type in the app's lock screen
 *   CLAUDE_API_KEY  = your Anthropic API key (for the AI chat)
 * Every request must include "key" = APP_KEY.
 */

function prop_(name) { return PropertiesService.getScriptProperties().getProperty(name) || ""; }
const DRIVE_FOLDER_NAME = "PeopleAppPhotos"; // created automatically in your Drive

// ---------- Sheet helpers ----------

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    return sheet;
  }
  // Add any header that an older sheet is missing (e.g. "color" on Categories)
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  headers.forEach((h) => {
    if (current.indexOf(h) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      current.push(h);
    }
  });
  return sheet;
}

// Column number (1-based) of a header, read from the sheet itself
function colOf_(sheet, header) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(header) + 1;
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter((row) => row[0] !== "" && row[0] !== null)
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function appendObject_(sheet, headers, obj) {
  const real = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = real.map((h) => (obj[h] !== undefined ? obj[h] : ""));
  sheet.appendRow(row);
}

function findRowIndexById_(sheet, idColName, id) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(idColName);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1; // 1-indexed for setRange
  }
  return -1;
}

function newId_(prefix) {
  return prefix + "_" + Utilities.getUuid().slice(0, 8);
}

// ---------- Sheet definitions ----------

const PEOPLE_HEADERS = ["id", "name", "categories", "email", "birthDate", "imageUrl", "featured"];
const RELATIONSHIPS_HEADERS = ["id", "personA", "personB", "type"];
const ABOUT_HEADERS = ["id", "personId", "date", "text"];
const FIELD_DEFS_HEADERS = ["id", "label", "type", "important"];
const FIELD_VALUES_HEADERS = ["id", "personId", "fieldId", "value"];
const CATEGORIES_HEADERS = ["name", "imageUrl", "color"];
const ITEMS_HEADERS = ["id", "collection", "name", "fields"];

function peopleSheet_() { return getSheet_("People", PEOPLE_HEADERS); }
function relationshipsSheet_() { return getSheet_("Relationships", RELATIONSHIPS_HEADERS); }
function aboutSheet_() { return getSheet_("About", ABOUT_HEADERS); }
function fieldDefsSheet_() { return getSheet_("FieldDefinitions", FIELD_DEFS_HEADERS); }
function fieldValuesSheet_() { return getSheet_("FieldValues", FIELD_VALUES_HEADERS); }
function categoriesSheet_() { return getSheet_("Categories", CATEGORIES_HEADERS); }
function itemsSheet_() { return getSheet_("Items", ITEMS_HEADERS); }

function parseFields_(v) {
  if (!v) return {};
  if (typeof v === "object") return v;
  try { return JSON.parse(v); } catch (e) { return {}; }
}

// ---------- Public API ----------
// All actions are routed through doGet/doPost below.
// Each function returns a plain JS object; doGet/doPost wraps it as JSON.

function getAllData_() {
  const people = sheetToObjects_(peopleSheet_()).map((p) => ({
    ...p,
    categories: p.categories ? String(p.categories).split(",").map((c) => c.trim()).filter(Boolean) : [],
    featured: p.featured === true || p.featured === "TRUE" || p.featured === "true",
  }));
  const relationships = sheetToObjects_(relationshipsSheet_());
  const aboutRows = sheetToObjects_(aboutSheet_());
  const fieldDefs = sheetToObjects_(fieldDefsSheet_()).map((f) => ({ ...f, important: f.important === true || f.important === "TRUE" || f.important === "true" }));
  const fieldValueRows = sheetToObjects_(fieldValuesSheet_());
  const categories = sheetToObjects_(categoriesSheet_());

  // Group About entries by personId
  const aboutByPerson = {};
  aboutRows.forEach((row) => {
    if (!aboutByPerson[row.personId]) aboutByPerson[row.personId] = [];
    aboutByPerson[row.personId].push({ date: row.date, text: row.text, id: row.id });
  });

  // Group FieldValues by personId -> fieldId -> value
  const valuesByPerson = {};
  fieldValueRows.forEach((row) => {
    if (!valuesByPerson[row.personId]) valuesByPerson[row.personId] = {};
    valuesByPerson[row.personId][row.fieldId] = row.value;
  });

  const peopleFull = people.map((p) => ({
    ...p,
    about: aboutByPerson[p.id] || [],
    customValues: valuesByPerson[p.id] || {},
  }));

  const items = sheetToObjects_(itemsSheet_()).map((it) => ({
    id: it.id,
    collection: it.collection,
    name: it.name,
    fields: parseFields_(it.fields),
    about: aboutByPerson[it.id] || [],
  }));

  return {
    people: peopleFull,
    relationships,
    fieldDefs,
    categories,
    items,
  };
}

function addPerson_(params) {
  const id = newId_("p");
  const sheet = peopleSheet_();
  appendObject_(sheet, PEOPLE_HEADERS, {
    id,
    name: params.name || "",
    categories: Array.isArray(params.categories) ? params.categories.join(",") : (params.categories || ""),
    email: params.email || "",
    birthDate: params.birthDate || "",
    imageUrl: params.imageUrl || "",
  });
  return { id };
}

function updatePerson_(params) {
  const sheet = peopleSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.id);
  if (rowIndex === -1) return { error: "Person not found" };
  const headers = PEOPLE_HEADERS;
  const current = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const updated = headers.map((h, i) => {
    if (h === "id") return current[i];
    if (params[h] === undefined) return current[i];
    if (h === "categories" && Array.isArray(params[h])) return params[h].join(",");
    return params[h];
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([updated]);
  return { ok: true };
}

function addRelationship_(params) {
  const id = newId_("r");
  appendObject_(relationshipsSheet_(), RELATIONSHIPS_HEADERS, {
    id,
    personA: params.personA,
    personB: params.personB,
    type: params.type || "",
  });
  return { id };
}

// NEW: Remove a relationship row between two people (by personA/personB in either order)
function removeRelationship_(params) {
  const sheet = relationshipsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const aCol = headers.indexOf("personA");
  const bCol = headers.indexOf("personB");
  // Delete rows bottom-up to preserve indices
  for (let i = data.length - 1; i >= 1; i--) {
    const a = String(data[i][aCol]);
    const b = String(data[i][bCol]);
    const matchAB = a === String(params.personA) && b === String(params.personB);
    const matchBA = a === String(params.personB) && b === String(params.personA);
    if (matchAB || matchBA) {
      sheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// NEW: Remove a category from a person's categories column in the People sheet
function removePersonFromCategory_(params) {
  const sheet = peopleSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.personId);
  if (rowIndex === -1) return { error: "Person not found" };
  const catColIndex = PEOPLE_HEADERS.indexOf("categories") + 1;
  const current = sheet.getRange(rowIndex, catColIndex).getValue();
  const cats = current ? String(current).split(",").map(c => c.trim()).filter(Boolean) : [];
  const newCats = cats.filter(c => c !== params.categoryName);
  sheet.getRange(rowIndex, catColIndex).setValue(newCats.join(","));
  return { ok: true };
}

// NEW: Add a category to a person's categories column in the People sheet
function addPersonToCategory_(params) {
  const sheet = peopleSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.personId);
  if (rowIndex === -1) return { error: "Person not found" };
  const catColIndex = PEOPLE_HEADERS.indexOf("categories") + 1;
  const current = sheet.getRange(rowIndex, catColIndex).getValue();
  const cats = current ? String(current).split(",").map(c => c.trim()).filter(Boolean) : [];
  if (!cats.includes(params.categoryName)) {
    cats.push(params.categoryName);
    sheet.getRange(rowIndex, catColIndex).setValue(cats.join(","));
  }
  return { ok: true };
}

function addAbout_(params) {
  const id = newId_("a");
  appendObject_(aboutSheet_(), ABOUT_HEADERS, {
    id,
    personId: params.personId,
    date: params.date || new Date().toISOString().slice(0, 10),
    text: params.text || "",
  });
  return { id };
}

function updateAbout_(params) {
  const sheet = aboutSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.id);
  if (rowIndex === -1) return { error: "Note not found" };
  if (params.text !== undefined) sheet.getRange(rowIndex, colOf_(sheet, "text")).setValue(params.text);
  if (params.date !== undefined) sheet.getRange(rowIndex, colOf_(sheet, "date")).setValue(params.date);
  return { ok: true };
}

function deleteAbout_(params) {
  const sheet = aboutSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.id);
  if (rowIndex !== -1) sheet.deleteRow(rowIndex);
  return { ok: true };
}

// AI chat: the Anthropic key stays here on the server, never in the page
function askClaude_(params) {
  const apiKey = prop_("CLAUDE_API_KEY");
  if (!apiKey) return { error: "CLAUDE_API_KEY is not set in Script Properties" };
  const res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify({
      model: "claude-haiku-4-5",
      max_tokens: 1024,
      system: String(params.system || ""),
      messages: [{ role: "user", content: String(params.message || "") }],
    }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200) return { error: "Claude API error " + code };
  const data = JSON.parse(res.getContentText());
  return { text: (data.content && data.content[0] && data.content[0].text) || "" };
}

function addFieldDef_(params) {
  // Use the id the app created, so values saved right away still match after reload
  const id = params.id || newId_("f");
  appendObject_(fieldDefsSheet_(), FIELD_DEFS_HEADERS, {
    id,
    label: params.label || "",
    type: params.type || "text",
    important: params.important ? "TRUE" : "FALSE",
  });
  return { id };
}

function setFieldValue_(params) {
  const sheet = fieldValuesSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const personCol = headers.indexOf("personId");
  const fieldCol = headers.indexOf("fieldId");
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][personCol]) === String(params.personId) && String(data[i][fieldCol]) === String(params.fieldId)) {
      sheet.getRange(i + 1, headers.indexOf("value") + 1).setValue(params.value);
      return { ok: true };
    }
  }
  const id = newId_("fv");
  appendObject_(sheet, FIELD_VALUES_HEADERS, { id, personId: params.personId, fieldId: params.fieldId, value: params.value });
  return { ok: true };
}

function addCategory_(params) {
  appendObject_(categoriesSheet_(), CATEGORIES_HEADERS, { name: params.name, imageUrl: params.imageUrl || "", color: params.color || "" });
  return { ok: true };
}

function updateCategory_(params) {
  const sheet = categoriesSheet_();
  const rowIndex = findRowIndexById_(sheet, "name", params.id || params.name);
  if (rowIndex === -1) return { error: "Category not found" };
  if (params.color !== undefined) sheet.getRange(rowIndex, colOf_(sheet, "color")).setValue(params.color);
  return { ok: true };
}

function renameCategory_(params) {
  const catSheet = categoriesSheet_();
  const data = catSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.oldName) {
      catSheet.getRange(i + 1, 1).setValue(params.newName);
    }
  }
  const peopleSheetRef = peopleSheet_();
  const pdata = peopleSheetRef.getDataRange().getValues();
  const catColIndex = PEOPLE_HEADERS.indexOf("categories");
  for (let i = 1; i < pdata.length; i++) {
    const cats = String(pdata[i][catColIndex] || "").split(",").map((c) => c.trim());
    if (cats.includes(params.oldName)) {
      const newCats = cats.map((c) => (c === params.oldName ? params.newName : c));
      peopleSheetRef.getRange(i + 1, catColIndex + 1).setValue(newCats.join(","));
    }
  }
  // Sub-category links in Relationships use the category name too
  const relSheet = relationshipsSheet_();
  const rdata = relSheet.getDataRange().getValues();
  const aCol = rdata[0].indexOf("personA"), bCol = rdata[0].indexOf("personB");
  for (let i = 1; i < rdata.length; i++) {
    if (rdata[i][aCol] === params.oldName) relSheet.getRange(i + 1, aCol + 1).setValue(params.newName);
    if (rdata[i][bCol] === params.oldName) relSheet.getRange(i + 1, bCol + 1).setValue(params.newName);
  }
  return { ok: true };
}

function deletePerson_(params) {
  const sheet = peopleSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.id);
  if (rowIndex !== -1) sheet.deleteRow(rowIndex);
  // Also delete all relationships involving this person
  const relSheet = relationshipsSheet_();
  const data = relSheet.getDataRange().getValues();
  const headers = data[0];
  const aCol = headers.indexOf("personA");
  const bCol = headers.indexOf("personB");
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][aCol]) === String(params.id) || String(data[i][bCol]) === String(params.id)) {
      relSheet.deleteRow(i + 1);
    }
  }
  // Delete about entries
  const aboutSheetRef = aboutSheet_();
  const aData = aboutSheetRef.getDataRange().getValues();
  const aHeaders = aData[0];
  const pCol = aHeaders.indexOf("personId");
  for (let i = aData.length - 1; i >= 1; i--) {
    if (String(aData[i][pCol]) === String(params.id)) aboutSheetRef.deleteRow(i + 1);
  }
  return { ok: true };
}

function deleteCategory_(params) {
  const key = params.id || params.name;
  const catSheet = categoriesSheet_();
  const data = catSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === key) catSheet.deleteRow(i + 1);
  }
  const peopleSheetRef = peopleSheet_();
  const pdata = peopleSheetRef.getDataRange().getValues();
  const catColIndex = PEOPLE_HEADERS.indexOf("categories");
  for (let i = 1; i < pdata.length; i++) {
    const cats = String(pdata[i][catColIndex] || "").split(",").map((c) => c.trim()).filter(Boolean);
    if (cats.includes(key)) {
      const newCats = cats.filter((c) => c !== key);
      peopleSheetRef.getRange(i + 1, catColIndex + 1).setValue(newCats.join(","));
    }
  }
  // Also remove any relationships involving this category
  const relSheet = relationshipsSheet_();
  const rdata = relSheet.getDataRange().getValues();
  const raCol = rdata[0].indexOf("personA"), rbCol = rdata[0].indexOf("personB");
  for (let i = rdata.length - 1; i >= 1; i--) {
    if (rdata[i][raCol] === key || rdata[i][rbCol] === key) relSheet.deleteRow(i + 1);
  }
  return { ok: true };
}

// ---------- Collections (books / places / recipes) ----------

function addItem_(params) {
  const id = newId_("i");
  appendObject_(itemsSheet_(), ITEMS_HEADERS, {
    id,
    collection: params.collection || "",
    name: params.name || "",
    fields: JSON.stringify(parseFields_(params.fields)),
  });
  return { id };
}

function updateItem_(params) {
  const sheet = itemsSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.id);
  if (rowIndex === -1) return { error: "Item not found" };
  if (params.name !== undefined) sheet.getRange(rowIndex, ITEMS_HEADERS.indexOf("name") + 1).setValue(params.name);
  if (params.fields !== undefined) sheet.getRange(rowIndex, ITEMS_HEADERS.indexOf("fields") + 1).setValue(JSON.stringify(parseFields_(params.fields)));
  return { ok: true };
}

function deleteItem_(params) {
  const sheet = itemsSheet_();
  const rowIndex = findRowIndexById_(sheet, "id", params.id);
  if (rowIndex !== -1) sheet.deleteRow(rowIndex);
  // Remove "Recommend to" links and notes for this item
  const relSheet = relationshipsSheet_();
  const rdata = relSheet.getDataRange().getValues();
  const rh = rdata[0];
  const aCol = rh.indexOf("personA"), bCol = rh.indexOf("personB");
  for (let i = rdata.length - 1; i >= 1; i--) {
    if (String(rdata[i][aCol]) === String(params.id) || String(rdata[i][bCol]) === String(params.id)) relSheet.deleteRow(i + 1);
  }
  const aSheet = aboutSheet_();
  const adata = aSheet.getDataRange().getValues();
  const pCol = adata[0].indexOf("personId");
  for (let i = adata.length - 1; i >= 1; i--) {
    if (String(adata[i][pCol]) === String(params.id)) aSheet.deleteRow(i + 1);
  }
  return { ok: true };
}

/**
 * Uploads a base64-encoded image to a dedicated Drive folder and returns
 * a shareable URL. Used for both person avatars and category images.
 */
function uploadImage_(params) {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

  const base64 = params.imageData.split(",").pop(); // strip "data:image/png;base64,"
  const mimeMatch = params.imageData.match(/^data:(image\/[a-zA-Z]+);base64,/);
  const mimeType = mimeMatch ? mimeMatch[1] : "image/png";
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, params.fileName || "photo.png");

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const directUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
  return { imageUrl: directUrl, fileId: file.getId() };
}

// ---------- HTTP entry points ----------

function doGet(e) {
  return handleRequest_(e.parameter);
}

function doPost(e) {
  let params = {};
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ error: "Invalid JSON body" });
  }
  return handleRequest_(params);
}

function handleRequest_(params) {
  const appKey = prop_("APP_KEY");
  if (!appKey || String(params.key) !== appKey) {
    return jsonResponse_({ error: "Unauthorized" });
  }

  try {
    let result;
    switch (params.action) {
      case "ping":
        result = { ok: true };
        break;
      case "askClaude":
        result = askClaude_(params);
        break;
      case "updateAbout":
        result = updateAbout_(params);
        break;
      case "deleteAbout":
        result = deleteAbout_(params);
        break;
      case "updateCategory":
        result = updateCategory_(params);
        break;
      case "getAllData":
        result = getAllData_();
        break;
      case "addPerson":
        result = addPerson_(params);
        break;
      case "updatePerson":
        result = updatePerson_(params);
        break;
      case "addRelationship":
        result = addRelationship_(params);
        break;
      case "removeRelationship":
        result = removeRelationship_(params);
        break;
      case "removePersonFromCategory":
        result = removePersonFromCategory_(params);
        break;
      case "addPersonToCategory":
        result = addPersonToCategory_(params);
        break;
      case "addAbout":
        result = addAbout_(params);
        break;
      case "addFieldDef":
        result = addFieldDef_(params);
        break;
      case "setFieldValue":
        result = setFieldValue_(params);
        break;
      case "addCategory":
        result = addCategory_(params);
        break;
      case "renameCategory":
        result = renameCategory_(params);
        break;
      case "deletePerson":
        result = deletePerson_(params);
        break;
      case "deleteCategory":
        result = deleteCategory_(params);
        break;
      case "addItem":
        result = addItem_(params);
        break;
      case "updateItem":
        result = updateItem_(params);
        break;
      case "deleteItem":
        result = deleteItem_(params);
        break;
      case "uploadImage":
        result = uploadImage_(params);
        break;
      default:
        result = { error: "Unknown action: " + params.action };
    }
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ error: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

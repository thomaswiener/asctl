var CREDENTIALS = {
  private_key: "$PRIVATE_KEY",
  client_email: "app-script@poetic-glass-489408-n6.iam.gserviceaccount.com"
}
var CONFIG = {
  table_name: "$TABLE_NAME",
  spreadsheet_id: "$SPREADSHEET_ID",
  sheet_name: "SQL_DWH"
}

function getOAuthService(user) {
  return OAuth2.createService("Service Account")
    .setTokenUrl("https://oauth2.googleapis.com/token")
    .setPrivateKey(CREDENTIALS.private_key)
    .setIssuer(CREDENTIALS.client_email)
    .setSubject(user)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setParam("access_type", "offline")
    .setScope("https://www.googleapis.com/auth/bigquery https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/devstorage.read_only");
}

function reset() {
  const service = getOAuthService(CREDENTIALS.client_email);
  service.reset();
}

function executeQuery(query) {
  const service = getOAuthService(CREDENTIALS.client_email);

  if (!service.hasAccess()) {
    throw new Error("Authorization required. Check script permissions.");
  }

  const request = {
    query: query,
    useLegacySql: false
  };

  const projectId = "poetic-glass-489408-n6";
  return BigQuery.Jobs.query(request, projectId);
}

function processQueryResults(queryResults, sheet, headers, startRow) {
  if (!sheet) {
    throw new Error("Sheet object is null.");
  }

  if (queryResults.rows && queryResults.rows.length > 0) {
    const data = queryResults.rows.map(row => row.f.map(cell => cell.v));
    const lastRow = sheet.getLastRow();

    if (lastRow >= startRow) {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, headers.length).clearContent();
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(startRow, 1, data.length, headers.length).setValues(data);
    Logger.log(`Imported ${data.length} rows into sheet "${sheet.getName()}".`);
  } else {
    Logger.log("No data found.");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const lastRow = sheet.getLastRow();
    if (lastRow >= startRow) {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, headers.length).clearContent();
    }
  }
}

function FetchResults() {
  const table = CONFIG.table_name
  const spreadsheetId = CONFIG.spreadsheet_id
  const sheetName = CONFIG.sheet_name
  const startRow = 2;

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(sheetName);

  // Falls Blatt nicht existiert, automatisch anlegen
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found. Creating new sheet.`);
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const query = `
    SELECT *
    FROM \`poetic-glass-489408-n6.dataset001.${table}\`
  `;

  const queryResults = executeQuery(query);
  const headers = queryResults.schema.fields.map(field => field.name);
  processQueryResults(queryResults, sheet, headers, startRow);
}


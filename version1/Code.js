//this code is for the Equipoise+ v1.18 GHG Calculator

function exiodropdown() {
  var activeCell=SpreadsheetApp.getActiveRange();
  var activeRow=activeCell.getRow()
  var activeCol=activeCell.getColumn()
  var activeValue=activeCell.getValue()
  var activeSheet=activeCell.getSheet()

  if(activeSheet.getName()=="Goods & Services" && activeRow>15 && activeCol>=6 && activeCol <=7){
    var worksheet=SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet=worksheet.getSheetByName("EXIOBASE")
    var data=spreadsheet.getDataRange().getValues();
    var list=data.filter(row=>row[activeCol-6]==activeValue).map(row=>row[activeCol-5])
    var validation =SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build()
    activeCell.offset(0,1).setDataValidation(validation)
  }
}

function fueldropdown() {
  var activeCell=SpreadsheetApp.getActiveRange();
  var activeRow=activeCell.getRow()
  var activeCol=activeCell.getColumn()
  var activeValue=activeCell.getValue()
  var activeSheet=activeCell.getSheet()

  if((activeSheet.getName()=="Fuel" || activeSheet.getName()=="Fuel (MA)") && activeRow>8 && activeCol==6){
    var worksheet=SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet=worksheet.getSheetByName("FuelsDropdowns")
    var data=spreadsheet.getDataRange().getValues();
    var list=data.filter(row=>row[activeCol-6]==activeValue).map(row=>row[activeCol-5])
    var validation =SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build()
    activeCell.offset(0,1).setDataValidation(validation)
  }
}

function introdropdown() {
  var activeCell=SpreadsheetApp.getActiveRange();
  var activeRow=activeCell.getRow()
  var activeCol=activeCell.getColumn()
  var activeValue=activeCell.getValue()
  var activeSheet=activeCell.getSheet()

  if(activeSheet.getName()=="Intro" && (activeRow>34 && activeRow<37) && activeCol==7){
    var worksheet=SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet=worksheet.getSheetByName("FuelsDropdowns")
    var data=spreadsheet.getDataRange().getValues();
    var list=data.filter(row=>row[activeCol-7]==activeValue).map(row=>row[activeCol-6])
    var validation =SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build()
    activeCell.offset(0,1).setDataValidation(validation)
  }
}

function transportdropdown() {
  var activeCell=SpreadsheetApp.getActiveRange();
  var activeRow=activeCell.getRow()
  var activeCol=activeCell.getColumn()
  var activeValue=activeCell.getValue()
  var activeSheet=activeCell.getSheet()

  if((activeSheet.getName()=="Inbound T&D" || activeSheet.getName()=="Outbound T&D") && activeRow>10 && activeCol==4){
    var worksheet=SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet=worksheet.getSheetByName("TransportDropdowns")
    var data=spreadsheet.getDataRange().getValues();
    var list=data.filter(row=>row[activeCol-4]==activeValue).map(row=>row[activeCol-3])
    var validation =SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build()
    activeCell.offset(0,1).setDataValidation(validation)
  }
}

function traveldropdown() {
  var activeCell=SpreadsheetApp.getActiveRange();
  var activeRow=activeCell.getRow()
  var activeCol=activeCell.getColumn()
  var activeValue=activeCell.getValue()
  var activeSheet=activeCell.getSheet()

  if((activeSheet.getName()=="Business Travel" || activeSheet.getName()=="Commuting") && activeRow>10 && activeCol==4){
    var worksheet=SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet=worksheet.getSheetByName("TravelDropdowns")
    var data=spreadsheet.getDataRange().getValues();
    var list=data.filter(row=>row[activeCol-4]==activeValue).map(row=>row[activeCol-3])
    var validation =SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build()
    activeCell.offset(0,1).setDataValidation(validation)
  }
}

function wastedropdown() {
  var activeCell=SpreadsheetApp.getActiveRange();
  var activeRow=activeCell.getRow()
  var activeCol=activeCell.getColumn()
  var activeValue=activeCell.getValue()
  var activeSheet=activeCell.getSheet()

  if(activeSheet.getName()=="Waste" && activeRow>10 && activeCol>=5 && activeCol <=6){
    var worksheet=SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet=worksheet.getSheetByName("WasteDropdowns")
    var data=spreadsheet.getDataRange().getValues();
    var list=data.filter(row=>row[activeCol-5]==activeValue).map(row=>row[activeCol-4])
    var validation =SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build()
    activeCell.offset(0,1).setDataValidation(validation)
  }
}

function griddropdown() {
  var activeCell=SpreadsheetApp.getActiveRange();
  var activeRow=activeCell.getRow()
  var activeCol=activeCell.getColumn()
  var activeValue=activeCell.getValue()
  var activeSheet=activeCell.getSheet()

  if(activeSheet.getName()=="Info" && activeRow>28 && activeCol==6){
    var worksheet=SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheet=worksheet.getSheetByName("GridDropdowns")
    var data=spreadsheet.getDataRange().getValues();
    var list=data.filter(row=>row[activeCol-6]==activeValue).map(row=>row[activeCol-5])
    var validation =SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(false).build()
    activeCell.offset(0,1).setDataValidation(validation)
  }
}

function onEdit() {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  exiodropdown();
  introdropdown();
  fueldropdown();
  transportdropdown();
  traveldropdown();
  wastedropdown();
  griddropdown();

  if (sheet.getName() === "Accom" && range.getColumn() === 5) {
    if (getApiToggleStatus("postAccomDataToAPI")) {
      const result = postAccomDataToAPI(/* required parameters */); // Ensure this function is correctly parameterized
      range.offset(0, 1).setValue(result); // Update adjacent cell with API result
    } else {
      range.offset(0, 1).setValue("API OFF"); // Show warning
    }
  }

  PropertiesService.getDocumentProperties().setProperty("lastActivityTime", new Date().getTime());
  ensureAutoTurnOffTrigger();
}


function onOpen() {
  updateMenu();
  ensureAutoTurnOffTrigger();
}

function updateMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Equipoise');
  menu.addItem(`Toggle All API Calls ON`, 'toggleAllAPIsOn');
  menu.addItem(`Toggle All API Calls OFF`, 'toggleAllAPIsOff');

  const apiFunctions = [
    "postDataToAPI",
    "postFreightDataToAPI",
    "postTravelDataToAPI",
    "postAccomDataToAPI",
    "postSpendDataToAPI"
  ];
  
  apiFunctions.forEach(apiFunction => {
    const formattedName = apiFunction
      .replace(/([A-Z])/g, ' $1')
      .replace(/\bA P I\b/, 'API');
    menu.addItem(
      `Toggle ${formattedName} ${getStatusText(apiFunction)}`,
      `toggle${apiFunction.charAt(0).toUpperCase() + apiFunction.slice(1)}`
    );
  });

  menu.addItem("Debug API Statuses", "debugApiStatuses"); // Add a debugging menu item
  menu.addToUi();
}

function toggleAllAPIs(on) {
  const apiFunctions = [
    "postDataToAPI",
    "postFreightDataToAPI",
    "postTravelDataToAPI",
    "postAccomDataToAPI",
    "postSpendDataToAPI"
  ];
  apiFunctions.forEach(apiFunction => setApiToggleStatus(apiFunction, on));
  updateMenu();
}

function toggleAllAPIsOn() {
  toggleAllAPIs(true);
}

function toggleAllAPIsOff() {
  toggleAllAPIs(false);
}

function autoTurnOffAPIs() {
  const lastActivityTime = PropertiesService.getDocumentProperties().getProperty("lastActivityTime");
  const now = new Date().getTime();
  const idleTimeout = 60000; // Adjust to 3600000 (1 hour) for production

  if (lastActivityTime && now - parseInt(lastActivityTime, 10) > idleTimeout) {
    toggleAllAPIs(false); // Automatically turn off all APIs
  }
}

function ensureAutoTurnOffTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerExists = triggers.some(trigger => trigger.getHandlerFunction() === "autoTurnOffAPIs");

  if (!triggerExists) {
    ScriptApp.newTrigger("autoTurnOffAPIs")
      .timeBased()
      .everyMinutes(1) // For testing; change to `.everyHours(1)` later
      .create();
  }
}

function debugApiStatuses() {
  const statuses = [
    "postDataToAPI",
    "postFreightDataToAPI",
    "postTravelDataToAPI",
    "postAccomDataToAPI",
    "postSpendDataToAPI"
  ].map(apiFunction => {
    return `${apiFunction}: ${getApiToggleStatus(apiFunction) ? "ON" : "OFF"}`;
  });

  SpreadsheetApp.getUi().alert("API Statuses:\n" + statuses.join("\n"));
}

function getStatusText(apiFunction) {
  return getApiToggleStatus(apiFunction) ? "(ON)" : "(OFF)";
}

function getApiToggleStatus(apiFunction) {
  return PropertiesService.getScriptProperties().getProperty(apiFunction) === "true";
}

function setApiToggleStatus(apiFunction, status) {
  PropertiesService.getScriptProperties().setProperty(apiFunction, status ? "true" : "false");
}

// Individual API toggle functions
function togglePostDataToAPI() {
  toggleSpecificAPI("postDataToAPI");
}

function togglePostFreightDataToAPI() {
  toggleSpecificAPI("postFreightDataToAPI");
}

function togglePostTravelDataToAPI() {
  toggleSpecificAPI("postTravelDataToAPI");
}

function togglePostAccomDataToAPI() {
  toggleSpecificAPI("postAccomDataToAPI");
}

function togglePostSpendDataToAPI() {
  toggleSpecificAPI("postSpendDataToAPI");
}

// Helper to toggle specific APIs
function toggleSpecificAPI(apiFunction) {
  const currentMode = getApiToggleStatus(apiFunction);
  const newMode = !currentMode;
  setApiToggleStatus(apiFunction, newMode);
  SpreadsheetApp.getUi().alert(`${apiFunction} is now ${newMode ? "ON" : "OFF"}`);
  updateMenu(); // Refresh the menu with updated status
}




function incrementApiCallCount() {
  var properties = PropertiesService.getScriptProperties();
  var currentCount = parseInt(properties.getProperty("apiCallCount")) || 0;
  properties.setProperty("apiCallCount", currentCount + 1); // Increment and store count
  return currentCount + 1;
}

function postDataToAPI(apiKeyCell, activityId, dataVersion, call_year, callRegion, parameter, value, unit, type, lca_activity) {
  const cell = SpreadsheetApp.getActiveRange();
  const cellKey = `postDataToAPI_${cell.getA1Notation()}`;

  if (!getApiToggleStatus("postDataToAPI")) {
    const lastValue = PropertiesService.getDocumentProperties().getProperty(cellKey);
    return lastValue ? lastValue : "API is OFF";
  }

  incrementApiCallCount();

  const apiUrl = "https://beta4.api.climatiq.io/estimate";
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getRange(apiKeyCell).getValue();
  const dataVersionValue = SpreadsheetApp.getActiveSpreadsheet().getRange(dataVersion).getValue();

  const requestData = {
    emission_factor: {
      activity_id: activityId,
      data_version: dataVersionValue,
      region: callRegion,
      year: call_year,
      source_lca_activity: lca_activity
    },
    parameters: {
      [parameter]: value,
      [parameter + '_unit']: unit
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(requestData)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());

    const result = type === "CO2e" ? data.co2e : type === "Source" ? data.emission_factor.source : "Invalid type specified";
    PropertiesService.getDocumentProperties().setProperty(cellKey, result);
    return result;

  } catch (error) {
    return "Error: " + error.message;
  }
}

function postFreightDataToAPI(apiKeyCell, activityId, dataVersion, call_year, callRegion, parameter, value, unit, d_unit, type, lca_activity) {
  const cell = SpreadsheetApp.getActiveRange();
  const cellKey = `postFreightDataToAPI_${cell.getA1Notation()}`;

  if (!getApiToggleStatus("postFreightDataToAPI")) {
    const lastValue = PropertiesService.getDocumentProperties().getProperty(cellKey);
    return lastValue ? lastValue : "API is OFF";
  }

  incrementApiCallCount();

  const apiUrl = "https://beta4.api.climatiq.io/estimate";
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getRange(apiKeyCell).getValue();
  const dataVersionValue = SpreadsheetApp.getActiveSpreadsheet().getRange(dataVersion).getValue();

  const requestData = {
    emission_factor: {
      activity_id: activityId,
      data_version: dataVersionValue,
      region: callRegion,
      year: call_year,
      source_lca_activity: lca_activity
    },
    parameters: {
      [parameter]: value,
      [parameter + '_unit']: unit,
      distance: 1,
      distance_unit: d_unit
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(requestData)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());

    const result = type === "CO2e" ? data.co2e : type === "Source" ? data.emission_factor.source : "Invalid type specified";
    PropertiesService.getDocumentProperties().setProperty(cellKey, result);
    return result;

  } catch (error) {
    return "Error: " + error.message;
  }
}


function postTravelDataToAPI(apiKeyCell, activityId, dataVersion, call_year, callRegion, parameter, value, d_unit, type, lca_activity) {
  const cell = SpreadsheetApp.getActiveRange();
  const cellKey = `postTravelDataToAPI_${cell.getA1Notation()}`;

  if (!getApiToggleStatus("postTravelDataToAPI")) {
    const lastValue = PropertiesService.getDocumentProperties().getProperty(cellKey);
    return lastValue ? lastValue : "API is OFF";
  }

  incrementApiCallCount();

  const apiUrl = "https://beta4.api.climatiq.io/estimate";
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getRange(apiKeyCell).getValue();
  const dataVersionValue = SpreadsheetApp.getActiveSpreadsheet().getRange(dataVersion).getValue();

  const requestData = {
    emission_factor: {
      activity_id: activityId,
      data_version: dataVersionValue,
      year: call_year,
      source_lca_activity: lca_activity
    },
    parameters: {
      [parameter]: value,
      distance: 1,
      distance_unit: d_unit
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(requestData)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());

    const result = type === "CO2e" ? data.co2e : type === "Source" ? data.emission_factor.source : "Invalid type specified";
    PropertiesService.getDocumentProperties().setProperty(cellKey, result);
    return result;

  } catch (error) {
    return "Error: " + error.message;
  }
}

function postAccomDataToAPI(apiKeyCell, activityId, dataVersion, call_year, callRegion, parameter, value, type, lca_activity) {
  const cell = SpreadsheetApp.getActiveRange();
  const cellKey = `postAccomDataToAPI_${cell.getA1Notation()}`;

  if (!getApiToggleStatus("postAccomDataToAPI")) {
    const lastValue = PropertiesService.getDocumentProperties().getProperty(cellKey);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "API is currently off. The value in this cell has not been updated.",
      "API Off",
      3 // Duration of 3 seconds
    );
    return lastValue ? lastValue : "API is OFF";
  }

  incrementApiCallCount();

  // Perform the API call
  const apiUrl = "https://beta4.api.climatiq.io/estimate";
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getRange(apiKeyCell).getValue();
  const dataVersionValue = SpreadsheetApp.getActiveSpreadsheet().getRange(dataVersion).getValue();

  const requestData = {
    emission_factor: {
      activity_id: activityId,
      data_version: dataVersionValue,
      region: callRegion,
      year: call_year,
      source_lca_activity: lca_activity
    },
    parameters: {
      [parameter]: value,
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(requestData)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());

    const result = type === "CO2e" ? data.co2e : type === "Source" ? data.emission_factor.source : "Invalid type specified";
    PropertiesService.getDocumentProperties().setProperty(cellKey, result);
    
    return result;

  } catch (error) {
    return "Error: " + error.message;
  }
}

function postingStale() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
      "API is currently off. The value in this cell has not been updated.",
      "API Off",
      3 // Duration of 3 seconds
    );
}

function postSpendDataToAPI(apiKeyCell, activityId, dataVersion, call_year, callRegion, value, unit, type, transport) {
  const cell = SpreadsheetApp.getActiveRange();
  const cellKey = `postSpendDataToAPI_${cell.getA1Notation()}`;

  if (!getApiToggleStatus("postSpendDataToAPI")) {
    
    postingStale();
    const lastValue = PropertiesService.getDocumentProperties().getProperty(cellKey);
    return lastValue ? lastValue : "API is OFF";
    
  }

  incrementApiCallCount();

  const apiUrl = "https://beta4.api.climatiq.io/procurement/spend";
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getRange(apiKeyCell).getValue();
  const dataVersionValue = SpreadsheetApp.getActiveSpreadsheet().getRange(dataVersion).getValue();

  const requestData = {
    activity: { activity_id: activityId },
    spend_year: call_year,
    spend_region: callRegion,
    money: value,
    money_unit: unit,
    transport_margin: transport
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(requestData)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());

    const result = type === "CO2e" ? data.estimate.co2e : type === "Source" ? data.estimate.emission_factor.source : "Invalid type specified";
    PropertiesService.getDocumentProperties().setProperty(cellKey, result);
    return result;

  } catch (error) {
    return "Error: " + error.message;
  }
}

function getApiCallCount() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty("apiCallCount") || 0;
}



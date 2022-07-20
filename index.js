const LOOKUP_SPREADSHEET = SpreadsheetApp.openById("<Lookup Sheet ID here>");
const FINAL_SPREADSHEET = SpreadsheetApp.openById("<Final Sheet ID here>");

// ------------------------------------------------------------
/**
 * Google hooks
 */

function onSubmit(e) {
  var answers = e;
  var responses = FormApp.getActiveForm().getResponses();
  // Logger.log("Responses: %s", responses[responses.length - 1]);
  // Logger.log(responses[responses.length - 1])
  var lastResponses = responses[responses.length - 1].getItemResponses();
  // Logger.log(lastResponses);
  // var vendors = getQuickbooksVendors();
  // Logger.log("Vendors API Response: %s", typeof(vendors));
  // var vendors_obj = JSON.parse(vendors)
  responses = {}
  for (var i = 0; i < lastResponses.length; i++) {
    responses[lastResponses[i].getItem().getTitle()] = lastResponses[i].getResponse()
  }
  Logger.log(responses)

  var contractorId = getUserId(responses["Last Name"], responses["First Name"]);
  var rate = getContractorJobRates(contractorId, responses["Site"], responses["Job Work"]);
  var finalRow = fillFinalTimecardSheet(responses, contractorId, rate);
  Logger.log("Final Row: ", finalRow);
}

// ------------------------------------------------------------

/**
 * Google Apps
 */

/**
 * Gets the contractors id.
 * Throws an error if there is no last name or first name found.
 * Return an int.
 */
function getUserId(lastName, firstName) {
  var contractors = LOOKUP_SPREADSHEET.getSheetByName("Contractor Numbers");
  var lastNameIndexes = contractors.createTextFinder(lastName).findAll();
  var result = lastNameIndexes.map(r => ({row: r.getRow(), col: r.getColumn()}));
  if (result.length <= 0) {
    throw new Error(`Could not find the contractors last name of ${lastName}.`);
  }
  Logger.log(result);
  var firstNameCol = -1;
  var lastNameRow = -1;
  for (var i = 0; i < lastNameIndexes.length; i++) {
    foundFirstName = contractors.getRange(result[i].row, result[i].col+1).getValue();
    Logger.log(foundFirstName);
    if (foundFirstName == firstName) {
      firstNameCol = result[i].col+1;
      lastNameRow = result[i].row;
      break;
    }
  }
  if (firstNameCol == -1 || firstNameCol == -1) {
    throw new Error(`Could not find the contractors first name of ${firstName}.`);
  }
  return contractors.getRange(lastNameRow, firstNameCol+1).getValue();
}

/**
 * Gets the rate for the contractorId at the given site and the given job.
 * Throws an error if there is no contractorId with a site and job's rate found.
 * Returns an int.
 */
function getContractorJobRates(contractorId, site, job) {
  var rateSheet = LOOKUP_SPREADSHEET.getSheetByName("Rates");
  var contractorIndexes = rateSheet.createTextFinder(contractorId).findAll();
  var results = contractorIndexes.map(r => ({row: r.getRow(), col: r.getColumn()}));
  var rate = -1;
  for (var i = 0; i < results.length; i++) {
    var foundSite = rateSheet.getRange(results[i].row, results[i].col+1).getValue();
    if (foundSite == site) {
      var foundJob = rateSheet.getRange(results[i].row, results[i].col+2).getValue();
      if (foundJob == job) {
        rate = rateSheet.getRange(results[i].row, results[i].col+3).getValue();
      }
    }
  }
  if (rate == -1) {
    throw new Error(`Could not find a rate for ${contractorId} with site ${site} and job ${job}.`);
  }
  return rate;
}

/**
 * 
 */
function fillFinalTimecardSheet(responses, contractorId, rate) {
  var sheet = FINAL_SPREADSHEET.getSheetByName("Sheet1");
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, lastColumn).getValues();
  var lastEmptyRow = sheet.getLastRow() + 1;
  var finalRange = sheet.getRange(lastEmptyRow, 1, 1, lastColumn);
  // {
  //   "Last Name",
  //   "First Name",
  //   "Day of the Week",
  //   "Date",
  //   "RT Start",
  //   "RT End",
  //   "Break",
  //   "OT Start",
  //   "OT End",
  //   "Job to Cost",
  //   "Total ST",
  //   "Total OT",
  //   "Rate RT",
  //   "Rate OT",
  //   "Other Expense Description",
  //   "Expense",
  //   "GST"
  // }
  var values = [[
    responses["Last Name"],
    responses["First Name"],
    `=switch(WEEKDAY(${sheet.getRange(lastEmptyRow, 4).getCell(1, 1).getA1Notation()}),1,"Sunday",2,"Monday",3,"Tuesday",4,"Wednesday",5, "Thursday",6,"Friday",7,"Saturday")`,
    `=DATEVALUE("${responses["Date"]}")`,
    responses["Start Time"],
    responses["End Time"],
    responses["Break"],
    "",
    "",
    responses["Show"] + " " + responses["Booth #"],
    `=${sheet.getRange(lastEmptyRow, 6).getCell(1,1).getA1Notation()}-${sheet.getRange(lastEmptyRow, 5).getCell(1,1).getA1Notation()}`,
    `=${sheet.getRange(lastEmptyRow, 9).getCell(1,1).getA1Notation()}-${sheet.getRange(lastEmptyRow, 8).getCell(1,1).getA1Notation()}`,
    rate,
    "",
    responses["Other Expense Description"],
    responses["Other Expense Total"],
    responses["GST"]
  ]];
  finalRange.setValues(values);
  return values;
}

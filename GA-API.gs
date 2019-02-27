/* Spreadsheet settings and common values */

var SS = SpreadsheetApp.getActiveSpreadsheet();
var CONFIG = SS.getSheetByName('Configuration');
var OUTPUT = SS.getSheetByName('Output');
var VIEW_ID = CONFIG.getRange('C2').getValue();
var FREQ = CONFIG.getRange('C3').getValue();
var START_DATE = Utilities.formatDate(CONFIG.getRange('C4').getValue(),'Europe/London', 'yyyy-MM-dd');
var END_DATE = Utilities.formatDate(CONFIG.getRange('C5').getValue(),'Europe/London', 'yyyy-MM-dd');

/* Read query paramters from config sheet
   Execute each query in sequence
   Write results back to output sheet
*/


function updateReport() {  
  var queries = getQueries();
  var outputRow = [];
  
  // Loop through queries, get results row from API response, and add to outputRow array
  var result = null;
  for (query in queries) {
    result = getResults(queries[query]);
    if (result.reports[0].data.rows) {
      result = [
        result.reports[0].data.rows[0].getDimensions()[0], 
        result.reports[0].data.rows[0].getMetrics()[0].values[0]
      ];
      outputRow.push(result[1]);
    }
  }

  // Take reporting period from last report and add to beginning of outputRow array
  outputRow.unshift(result[0]); 

  outputToSheet(outputRow);
  OUTPUT.activate();
  return;
}


/* Read queries from Configuration sheet and return as an array */

function getQueries() {
  var startRow = 9;
  var startCol = 3;
  var numRows = CONFIG.getLastRow() - startRow +1;
  var numCols = CONFIG.getLastColumn() - startCol + 1;
  var queries = CONFIG.getRange(startRow, startCol, numRows, numCols).getValues();
  return(queries);
}


/* Query GA API and returns result in an array */

function getResults(query) {
  
  // Set GA view id if provided or use default
  if(query[0]) { var tableId = query[0]; } else { var tableId = VIEW_ID; }
  
  var metrics = [{ "expression": query[1] }];
  var optArgs = {};  
  
  // Add filters and segments to optional arguements if provided
  if(query[2]) optArgs['filters'] = query[2];
  if(query[3]) optArgs['segment'] = [{ "segmentId": query[3] }];
  if(query[4]) optArgs['useResourceQuotas'] = query[4] == true;
  
  // Set dimension value depending on weekly or monthly report setting
  switch(FREQ) {
    case 'Weekly':
      optArgs['dimensions'] = 'ga:isoYearIsoWeek';
      break;
    case 'Monthly':
      optArgs['dimensions'] = 'ga:yearMonth';
      break;
    case 'Daily':
      optArgs['dimensions'] = 'ga:date';
      break;
  }
    
  var dimensions = [{ "name": optArgs['dimensions'] }];
  
  if (optArgs['segment']) {
    dimensions.push({
      "name": "ga:segment"
    });
  }
  
  // Gets results from the API
  var results = AnalyticsReporting.Reports.batchGet({
    "reportRequests":[
      {
        "viewId": tableId,
        "dateRanges":[
          {
            "startDate": START_DATE,
            "endDate": END_DATE
          }],
        "metrics": metrics,
        "dimensions": dimensions,
        "filtersExpression": optArgs['filters'],
        "segments": optArgs['segment']
      }],
      "useResourceQuotas": optArgs['useResourceQuotas']
  });
  
  return(results);
}


/* Appends results array to output sheet as a new row  */

function outputToSheet(outputRow) {
  if (OUTPUT != null) {
    OUTPUT.appendRow(outputRow);    
  }
}

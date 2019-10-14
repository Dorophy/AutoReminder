/**
* Automatically send emails to OKR owners who haven't updated their OKRs within the past 7 days on Thursdays 6-7pm
* Automatically send emails to managers for the team update summary on Fridays 6-7am
* owner: jing wang (jiwang2@linkedin.com)
*/

var OverallSummaryAlias = 'jiwang2@linkedin.com'; // if wanna add more, seperate by ','

var TeamAndManagerEmail = {
  'Discovery': {
    'managerName': 'Brent',
    'email': 'bmckee@linkedin.com'
  },
 
  'Streaming': {
    'managerName': 'Clark',
    'email':'chaskins@linkedin.com'
  },
 
'Infra Storage': {
'managerName': 'Hello',
'email': 'jiwang2@linkedin.com'
},
'DBE': {
'managerName': 'Hello',
'email':'jiwang2@linkedin.com'
},
'Oracle & Ambry': {
'managerName': 'Hello',
'email':'jiwang2@linkedin.com'
},

'Offline': {
'managerName': 'Hello',
'email':'jiwang2@linkedin.com'
}

};

var TABS = Object.keys(TeamAndManagerEmail);
var TEAMS = ['Discovery', 'Streaming', 'Infra Storage', 'DBE', 'Oracle & Ambry', 'Offline'];

/*
* Send reminder email to owners of each unupdated OKR
*/
function sendReminderEmail() {
  for (var i = 0; i < TABS.length; i++) {      
    var data = getDataByTabName(TABS[i]);
    var meta = data[0];
   
    var values = data.slice(1); // get the subarray from 1 line.
    var columnNameAndIndex = parseMetaData(meta);
    sendReminderEmailExec(columnNameAndIndex, values, TABS[i]);
  }
}


/*
* Send roll up summary email
*/

function sendRollUpSummaryEmail() {
   var rollUpTalbeSummary = {};
   var rollUpChartSummary = {
      'Done' : 0,
      'OnTrack': 0,
      'Postponed': 0,
      'AtRisk': 0,
      'NotStarted': 0,
   };
   var rollUpCommittedSummary = {};
   var rollUpTabName = 'Data SRE Roll Up';
   var data = getDataByTabName(rollUpTabName);
   var meta = data[0];
   var values = data.slice(1);
   var columnNameAndIndex = parseMetaData(meta);
 
   for (i in values) {
      var row = values[i];
      // chekc empty
      if (isEmptyRow(row)) {
          continue;
      }
     
      var status = row[columnNameAndIndex['Status']];
      var statusTrim = status.replace(/\s+/g, ''); // replace space by ''
     
      if (!rollUpChartSummary[statusTrim]) {
         rollUpChartSummary[statusTrim] = 0;
      }
      rollUpChartSummary[statusTrim]++;
         
      var teamName = row[columnNameAndIndex['Team']].trim();
     
      var committed = row[columnNameAndIndex['Committed']];
     
     
      if (TEAMS.indexOf(teamName) != -1) {
         if (!rollUpTalbeSummary[teamName]) {
            rollUpTalbeSummary[teamName] = {
               'Done' : 0,
               'OnTrack': 0,
               'Postponed': 0,
               'AtRisk': 0,
               'NotStarted': 0,
            };
         }
         rollUpTalbeSummary[teamName][statusTrim]++;
         
         
         if (!rollUpCommittedSummary[teamName]) {
             rollUpCommittedSummary[teamName] = {
                 'updatedNum': 0,
                 'committedNum': 0
             };
         }
         
         if (statusTrim != 'Done') {
            rollUpCommittedSummary[teamName].committedNum ++;
         }
         
         var lastModifyDate = row[row.length - 1];
         if (statusTrim != 'Done' && !notUpdatedMoreThanThresholdInDays(lastModifyDate, 7)) {
            rollUpCommittedSummary[teamName].updatedNum ++;
         }
       
      }
     
   }
   var test = rollUpCommittedSummary;
   
   /*
    updatedNum: updatedNum,
      committedNum: committedNum
   */
   
   // send the rollUp Summary
   var email = OverallSummaryAlias;
   sendRollUpEmailExec(rollUpChartSummary, rollUpTalbeSummary, rollUpCommittedSummary, email);
}


function sendRollUpEmailExec(rollUpChartSummary, rollUpTalbeSummary, rollUpCommittedSummary, emails) {
    var data = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, "Type")
      .addColumn(Charts.ColumnType.NUMBER, "Count")  
      .addRow(["AtRisk", rollUpChartSummary['AtRisk']]) // red
      .addRow(["NotStarted", rollUpChartSummary['NotStarted']])  // gray
      .addRow(["Postponed", rollUpChartSummary['Postponed']]) // orange
      .addRow(["Done", rollUpChartSummary['Done']]) // blue
      .addRow(["OnTrack", rollUpChartSummary['OnTrack']]) // green
      .build();
     
  var chartImage = Charts.newPieChart()
                    .setTitle('OKR Roll up')
                    .setOption('pieSliceText', 'value-and-percentage')// show number and percentage
                    .setColors(['red', 'orange', 'gray', 'blue', 'green'])
                    .setDimensions(1000, 1000)
                    .setDataTable(data)
                    .build()
                    .getAs('image/jpeg'); //get chart as image
                   
   
   var htmlTable = composeRollUpHtmlTable(rollUpTalbeSummary);
   
   var committedTable = composeHtmlTable(rollUpCommittedSummary, "AllTeam");
   
   var subject = "OKR Roll up3";
   var html = "<p>Hi managers,</p>" +"Here is a summary of the OKR Updates" + " <b>" + htmlTable +"</b>" + "<p> this is a test </p>" + " <b>" + committedTable +"</b>";
 
   // send email out
   MailApp.sendEmail({
       to: emails,
       subject: subject,
       htmlBody: html,
       inlineImages: {
         chartImg: chartImage,
       }
   });
   
}

function sendOKRSummaryChartExec(summary, allEmails){
   var data = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, "Type")
      .addColumn(Charts.ColumnType.NUMBER, "Count")  
      .addRow(["AtRisk", summary['AtRisk']]) // red
      .addRow(["Postponed", summary['Postponed']]) // orange
      .addRow(["NotStarted", summary['NotStarted']])  // gray
      .addRow(["Done", summary['Done']]) // blue
      .addRow(["OnTrack", summary['OnTrack']]) // green
      .build();
     
  var chartImage = Charts.newPieChart()
                    .setTitle('OKR Summary')
                    .setOption('pieSliceText', 'value-and-percentage')// show number and percentage
                    .setColors(['red', 'orange', 'gray', 'blue', 'green'])
                    .setDimensions(1000, 1000)
                    .setDataTable(data)
                    .build()
                    .getAs('image/jpeg'); //get chart as image

   for (i in allEmails) {
       var emailAddress = allEmails[i];
       MailApp.sendEmail({
       to: emailAddress,
       subject: "OKR summary",
       htmlBody: "<br> Please see the chart below for the latest OKR summary <img src='cid:chartImg'> <br>",
       inlineImages: {
         chartImg: chartImage,
       }
    });
  }  
}

function isEmptyRow(row) {
   for (var elem in row) {
      if (elem && elem.trim() !== '') {
            return false;
      }
   }
   return true;
}


/**
* Send OKR summary to alias
*/
function sendUpdateSummaryEmail() {
  var updateSummary = {};
 
  var updateSummaryByTeam = {};
 
  for (var i = 0; i < TABS.length; i++) {
    var tabName = TABS[i];
    var data = getDataByTabName(tabName);
    var meta = data[0];
    var values = data.slice(1);
    // parse meta
    var columnNameAndIndex = parseMetaData(meta);
    var updatedInfo = getUpdateSummaryPerTab(columnNameAndIndex, values, tabName);
    updateSummary[tabName] = updatedInfo.summaryInfo;
    updateSummaryByTeam[tabName] = updatedInfo.summaryInfoWithDetails;
  }
 
  /*
  * updateSummary
  {
      "Discovery":
      {
          updatedNum: 10,
          committedNum: 20
      },
      "Infra Storage":
      {
          updatedNum: 5,
          committedNum: 10
      },
      ...
  }
  */
  sendUpdateSummaryEmailOverallExec(updateSummary);
  sendUpdateSummaryEmailByTeamExec(updateSummaryByTeam);
 
}



/*
* Get data based on table name
*/
function getDataByTabName(tabName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(tabName);
  var startRow = 2; // First row of data to process
  var startColumn = 1; // status column
 
  var lastColumn = sheet.getLastColumn(); // last column index
  var lastRow = sheet.getLastRow();  // last row index
 
  var numRows = lastRow - startRow + 1;
  var numColumns = lastColumn - startColumn + 1;
 
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  var data = dataRange.getValues();
  return data;
}

/*
* Helper function of sendReminderEmail
*/
function sendReminderEmailExec(columnNameAndIndex, data, tabName) {
  var ownersAndOKRs = {};
  for (i in data) {
    // parse each row
    var row = data[i];
    var status = row[columnNameAndIndex['Status']];
    var OKRDescription = row[columnNameAndIndex['Project/Initative']];
    var owners = row[columnNameAndIndex['Owner']];
   
    var statusTrim = status.replace(/\s+/g, ''); // replace space by ''
    var committed = row[columnNameAndIndex['Committed']];
   
    var lastModifyDate = row[row.length - 1];
    var needToSendNoficiation = notUpdatedMoreThanThresholdInDays(lastModifyDate, 7) && status != 'Done' && committed === 'Yes';
   
    if (needToSendNoficiation) {
      owners = owners.split(',');  // "a, b, c" => ["a", " b", ]
     
      for (oIdx in owners) {
        var owner = owners[oIdx].replace(/\s+/g, '');
        var ownerEmail = owner + '@linkedin.com';
       
        if (!ownersAndOKRs[ownerEmail]) {
          ownersAndOKRs[ownerEmail] = [];
        }
        //push the okr to related owner's array as value
        ownersAndOKRs[ownerEmail].push(OKRDescription);
      }
    }
  }
 
  //sendOKRRequiredAction(ownersAndOKRs);
}


/*
* Return column name and its index
* Return example:
* {
*   "Rollup For Kevin?": 0,
*   "Mark BLR or US": 1,
*    ...
* }
*/
function parseMetaData(metaData) {
  var columnNameIndex = {};
  for (var idx = 0; idx < metaData.length; idx ++) {
    var cName = metaData[idx];
    columnNameIndex[cName] = idx;
  }
  return columnNameIndex;
}

/*
* Traverse one tag and get the updated OKR number and committed OKR number
* Return example
* {
*    updatedNum: 10,
*    committedNum: 20
* }
*/
function getUpdateSummaryPerTab(columnNameAndIndex, data, tabName) {
  var updatedNum = 0;
  var committedNum = 0;
  var infoByOwner = {};
  for (i in data) {
    var row = data[i];
   
    var owners = row[columnNameAndIndex['Owner']];
    owners = owners.split(',');  // "a, b, c" => ["a", " b", ]  
    var status = row[columnNameAndIndex['Status']];
    var statusTrim = status.replace(/\s+/g, ''); // replace space by ''
   
    var committed = row[columnNameAndIndex['Committed']];
    if (committed === 'Yes' && statusTrim != 'Done') {
      committedNum++;
     
      for (ownerIdx in owners) {
        var owner = owners[ownerIdx].replace(/\s+/g, '');
        if (owner === '') {
          owner = 'No-Owner';
        }
        if (!infoByOwner[owner]) {
          infoByOwner[owner] = {
            "updatedNum": 0,
            "committedNum": 0
          };
        }
        infoByOwner[owner].committedNum++;
      }
     
     
    }
   
    var lastModifyDate = row[row.length - 1];
    if (statusTrim != 'Done' && committed === 'Yes' && !notUpdatedMoreThanThresholdInDays(lastModifyDate, 7)) {
      updatedNum++;
      for (ownerIdx in owners) {
        var owner = owners[ownerIdx].replace(/\s+/g, '');
        if (infoByOwner[owner] && infoByOwner[owner].updatedNum) {
           infoByOwner[owner].updatedNum++;
        }
       
      }
    }
   
  }
 
  return {
    summaryInfo: {
      updatedNum: updatedNum,
      committedNum: committedNum
    },
    summaryInfoWithDetails: infoByOwner
  };
}


/*
 * Helper function of sendUpdateSummaryEmail
 */
function sendUpdateSummaryEmailOverallExec(updateSummary) {
  // create table html
  var htmlTable = composeHtmlTable(updateSummary, 'AllTeam');
 
  var subject = 'OKR Summary';
  var html = "<p>Hi managers,</p>" +"Here is a summary of the OKR Updates" + " <b>" + htmlTable +"</b>" +
    "<p>Thanks,<p>" +
      "<p>Jing<p>";
 
  // send email out
  var email = OverallSummaryAlias;
  try {
     MailApp.sendEmail({
       to: email,
       subject: subject,
       htmlBody: html
     });

  } catch(e) {
     console.log("Something is wrong with sendEmail");
  }

}


/*
 * Send OKRs update info for each team
 */
function sendUpdateSummaryEmailByTeamExec(updateSummaryByTeam) {
  if (updateSummaryByTeam) {
    for (var teamName in updateSummaryByTeam) {
      var summary = updateSummaryByTeam[teamName];
      var htmlTable = composeHtmlTable(summary, 'EachTeam');
      var email = TeamAndManagerEmail[teamName].email;
      var managerName = TeamAndManagerEmail[teamName].managerName;
     
      var subject = teamName + ' OKR Summary';
      var html = "<p>Hi " + managerName + ",</p>" +"Here is a summary of the " + teamName + " OKR Updates" + " <b>" + htmlTable +"</b>" +
        "<p>Thanks,</p>" +
          "<p>Jing</p>";
     
      // send email out
     
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: html
        });
      } catch (e) {
         console.log("Something is wrong with sendEmail");
      }
     
    }
  }
}



/*
 * Compose the OKR update summary html.
 * kind == AllTeam: compose the summary of all OKRs
 * kind == EachTeam: compose the summary of OKRs for one specific team
 */
function composeHtmlTable(updateSummary, kind) {
  var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  var htmltable = '<table ' + TABLEFORMAT +' ">';
 
  htmltable += '<tr>';
  if (kind === 'AllTeam') {
    htmltable += '<th>' + 'Team' + '</th>';
  } else {
    htmltable += '<th>' + 'Owner' + '</th>';
  }
  htmltable += '<th>' + 'Not-Updated' + '</th>';
  htmltable += '<th>' + 'Committed' + '</th>';
  htmltable += '<th>' + 'Not-Updated Percentage' + '</th>';
  htmltable += '</tr>';
 
  var totalUnUpdated = 0;
  var totalCommitted = 0;
 
  for (var tab in updateSummary) {
    var summary = updateSummary[tab];
    var updated = summary.updatedNum;
    var committed = summary.committedNum;
    var unUpdated = committed - updated;
   
    totalUnUpdated += unUpdated;
    totalCommitted += committed;
   
    htmltable += '<tr>';
    htmltable += '<td>' + tab + '</td>';
    htmltable += '<td>' + unUpdated + '</td>';
    htmltable += '<td>' + committed + '</td>';
    htmltable += '<td>' + parseFloat(unUpdated / committed * 100.00).toFixed(2) + '%' + '</td>';
    htmltable += '</tr>';
  }
 
  // add total info
  htmltable += '<tr>';
  htmltable += '<td>' + '<b>Total' + '</td>';
  htmltable += '<td>' + '<b>' + totalUnUpdated + '</td>';
  htmltable += '<td>' + '<b>' + totalCommitted + '</td>';
  htmltable += '<td>' + '<b>' + parseFloat(totalUnUpdated / totalCommitted * 100.00).toFixed(2) + '%' + '</td>';
  htmltable += '</tr>';
 
  htmltable += '</table>';
 
  return htmltable;
}


function composeRollUpHtmlTable(rollUpTableSummary) {
  var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  var htmltable = '<table ' + TABLEFORMAT +' ">';
 
  htmltable += '<tr>';
  var kind = null;
  htmltable += '<th>' + 'Team' + '</th>';
  htmltable += '<th>' + 'AtRisk' + '</th>';
  htmltable += '<th>' + 'OnTrack' + '</th>';
  htmltable += '<th>' + 'Postponed' + '</th>';
  htmltable += '<th>' + 'NotStarted' + '</th>';
  htmltable += '<th>' + 'Done' + '</th>';
  htmltable += '</tr>';
 
  for (var team in rollUpTableSummary) {
    var summary = rollUpTableSummary[team];
    var atRisk = summary['AtRisk'];
    var onTrack = summary['OnTrack'];
    var postponed = summary['Postponed'];
    var done = summary['Done'];
    var notStarted = summary['NotStarted'];
   
    htmltable += '<tr>';
    htmltable += '<td>' + team + '</td>';
    htmltable += '<td>' + atRisk + '</td>';
    htmltable += '<td>' + onTrack + '</td>';
    htmltable += '<td>' + postponed + '</td>';
    htmltable += '<td>' + notStarted + '</td>';
    htmltable += '<td>' + done + '</td>';
    htmltable += '</tr>';  
  }
 
  htmltable += '</table>';
  return htmltable;
}


/*
* Function send the unupdated OKR info to the owner of OKR
*/
function sendOKRRequiredAction(ownersAndOkrs) {
  for (var owner in ownersAndOkrs) {
    // [okr1,okr2]
    var okrsArray = ownersAndOkrs[owner];
    // combine the okrs into one string
    var okrsString = '';
    for (i in okrsArray) {
      if (okrsString === '') {
        okrsString = "<p>" + okrsArray[i];
      } else {
        okrsString = okrsString + " </p> <p> " + okrsArray[i];
      }  
    }
    okrsString = okrsString + "</p>";
   
    var subject = '[Action Required] OKR Items Update Needed';
    var html = "You are receiving this email as your OKR item(s)" + " <b>" + okrsString +"</b>" +
      "hasn’t been updated within the past 7 days" +
        "Please have it updated by Noon the upcoming Monday.<p>Here is the link: <a href=\"http://go/dataokrs\">http://go/dataokrs</a><p>" +
          "<p>Thanks,</p>" +
            "<p>Jing<p>";
    try {
      MailApp.sendEmail({
        to: owner,
        subject: subject,
        htmlBody: html
      });
   
    } catch(e) {
       
    }
   
  }
}


/*
* Get the date diff in days and return true if the diffInDays > thresholdInDays
*/
function notUpdatedMoreThanThresholdInDays(lastModifyDate, thresholdInDays) {
  var curDate = new Date();
 
  if (typeof lastModifyDate === 'string') {
    return false;
  }
  var lastModificationDate = lastModifyDate;
 
  var timeDiff = Math.abs(curDate.getTime() - lastModificationDate.getTime());
  var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24));
  if (diffDays > thresholdInDays) {
    return true;
  }
  return false;
}
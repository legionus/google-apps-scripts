/*
 * Copyright (C) 2018  Alexey Gladkov <gladkov.alexey@gmail.com>
 *
 * This file is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.

 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.

 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301, USA.
 */

/*
 * Documentation
 * https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
 * https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets
 */

var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1T9mZrWIUldc3DJvidUNHdyj_3nFdQaHbwzbwkPM7a7Q/edit');
var sheetName = 'Tests';

var mailLabel        = GmailApp.getUserLabelByName('jenkins');
var mailHandledLabel = GmailApp.getUserLabelByName('jenkins/spreadsheet');

var messageFailedSubjectRE = new RegExp("^(.*\])?[ ]*Build failed in Jenkins: ([^# ]+) #([0-9]+)", '');
var messageGoodSubjectRE   = new RegExp("^(.*\])?[ ]*Jenkins build is back to normal[ ]*:[ ]+([^#]+)[ ]+#([0-9]+)", '');
var messageUrlRE           = new RegExp('See <(https://[^>]+)>');

var weeks = 9;

function myFunction()
{
  if (mailLabel === null || mailHandledLabel === null) {
    Logger.log('No labels mailLabel=' + (mailLabel === null) + ', mailHandledLabel=' + (mailHandledLabel === null));
    return;
  }
  Logger.log(ss.getName());
  filterMails();
}

function hasLabel(thread, label)
{
  var labels = thread.getLabels();
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].getName() === label.getName()) {
      return true;
    }
  }
  return false;
}

function toType(obj)
{
  return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase()
}

/*
 * refreshFilterViews requires "Google Sheets API" advanced Google service.
 * https://developers.google.com/apps-script/guides/services/advanced?authuser=1#enabling_advanced_services
 */
function refreshFilterViews()
{
  var sheet = ss.getSheetByName(sheetName);
  var res = Sheets.Spreadsheets.get(ss.getId());

  var lastRow = sheet.getLastRow();
  var requests = [];

  for (var i = 0; i < res.sheets.length; i++) {
    if (res.sheets[i].filterViews !== undefined) {
      Logger.log("Update filter views for " + res.sheets[i].properties.title);
      for (var j = 0; j < res.sheets[i].filterViews.length; j++) {
        res.sheets[i].filterViews[j].range.endRowIndex = lastRow;
        requests.push({"updateFilterView": {"filter": res.sheets[i].filterViews[j], "fields": "*"}});
      }
    }
    if (res.sheets[i].basicFilter !== undefined) {
      Logger.log("Update basic filter for " + res.sheets[i].properties.title);
      res.sheets[i].basicFilter.range.endRowIndex = lastRow;
      requests.push({"setBasicFilter": {"filter": res.sheets[i].basicFilter}});
    }
  }
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({'requests': requests}, ss.getId());
  }
}

function processThread(sheet, thread)
{
  var message = thread.getMessages()[0];

  var subj = message.getSubject();
  var res, url, testStatus, testName, testJob;

  if ((res = messageFailedSubjectRE.exec(subj)) !== null) {
    testStatus = "Broken";
    testName = res[2];
    testJob = res[3];
  } else if ((res = messageGoodSubjectRE.exec(subj)) !== null) {
    testStatus = "Good";
    testName = res[2];
    testJob = res[3];
  } else {
    Logger.log('Unparsed:' + subj);
    return;
  }

  if (hasLabel(thread, mailHandledLabel)) {
    return;
  }
  thread.addLabel(mailHandledLabel);

  if ((res = messageUrlRE.exec(message.getPlainBody())) !== null) {
    url = '=HYPERLINK("' + res[1] + '", "CI LINK")';
  }

  var logURL = '=HYPERLINK("https://apps.dmage.ru/jenkins/' + testName + '/' + testJob + '", "LOG")';

  sheet.appendRow([message.getDate(), testStatus, testName, testJob, url, logURL]);
}

function filterMails()
{
  var sheet = ss.getSheetByName(sheetName);
  if (sheet === null) {
    Logger.log("filterJenkins: Sheet not found: " + sheetName);
    return;
  }

  var threads = GmailApp.search('label:' + mailLabel.getName() + ' -label:' + mailHandledLabel.getName(), 0, 100);

  var dirty = false;
  for (var i = 0; i < threads.length; i++) {
    processThread(sheet, threads[i]);
    dirty = true;
  }
  if (dirty) {
    refreshFilterViews();
  }
}

function removeObsoleted()
{
  var now = Date.now();

  var sheet = ss.getSheetByName(sheetName);
  if (sheet === null) {
    Logger.log("removeObsoleted: Sheet not found: " + sheetName);
    return;
  }

  var lastRow = sheet.getLastRow();

  for (var i = lastRow; i > 0; i--) {
    var range = sheet.getRange(i, 1);
    var values = range.getValues();
    var rowDate = new Date(values[0][0]);

    if (((now - rowDate)/1000) > (weeks * 604800)) {
      sheet.deleteRow(i);
    }
  }
}

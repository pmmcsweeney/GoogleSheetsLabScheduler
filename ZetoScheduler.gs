/* 
  Software: Zeto Machine Scheduling system
   Made by: Patrick McSweeney 
For use by: Western Regional Small Grains Genotyping Laboratory (WRSGGL)
*/


//on open function to initialize clickable UI elements
function onOpen() {
  var access = SpreadsheetApp.getActiveSheet().getRange(PropertiesService.getScriptProperties().getProperty('ACCESSTOGGLEROW'), PropertiesService.getScriptProperties().getProperty('ACCESSTOGGLECOL')).getValue();
  if (access == 1)
  { Browser.msgBox('Welcome to the Scheduling system for the Zeto Analyzer machine in Johnson Hall room 296. In order to properly use this sign up sheet, you MUST be logged in with a google account. If you need to add a plate please use the "Sign Up" button, or select the "Sign Up" menu item on top of the screen. Please do not alter the text in the "Scheduler" column. If you need to remove a plate, please use the "Remove Plate" option in the Sign Up section of the menu.');}
  else
  { var html = HtmlService.createHtmlOutput("<b><h1><font color=\"red\">ABI Sign Up Closed</font></b></h1>Do not schedule new plates at this time.<br><br> <input type=\"button\" value=\"Close\" onclick=\"google.script.host.close()\"  />").setWidth(500).setHeight(200).setSandboxMode(HtmlService.SandboxMode.IFRAME)
    SpreadsheetApp.getUi().showModalDialog(html, ' '); }
  /* ----------- This block handles what is displayed by menu items ------------- */
  
  SpreadsheetApp.getUi()
      .createMenu('Sign Up')
      .addItem('Sign Up', 'openDialog')
      .addItem('Remove Plate', 'RemovePlate')
      .addToUi();
  SpreadsheetApp.getUi()
      .createMenu('Admin')
      .addItem('Export and Remove Data', 'ExportAndRemoveMenuOption')
      .addItem('Auto-Correct Scheduler', 'CheckSum')
      .addItem('Toggle Zeto Access', 'ToggleZetoAccess')
      .addToUi();
  
  /* ------------- You must reload the spreadsheet to view changes -------------- */
  
  
  
  
  
  /* --------------- CONSTANTS: called by .getProperty("name") ------------------ */
  
  PropertiesService.getScriptProperties().setProperty('TIMECOLUMN', 13);
  PropertiesService.getScriptProperties().setProperty('SAVEID', '1cvDcj-RunRHxOfoTa68nY7bFMsNWpLBhzL6XYIGS0zg');
  PropertiesService.getScriptProperties().setProperty('STARTROW', 7);
  PropertiesService.getScriptProperties().setProperty('PLATETYPEROW', 8);
  PropertiesService.getScriptProperties().setProperty('ACCESSTOGGLECOL', 12);
  PropertiesService.getScriptProperties().setProperty('ACCESSTOGGLEROW', 2);
  PropertiesService.getScriptProperties().setProperty('PASSWORD', 'Deven3730');
  
  /* ------------- IF YOU ALTER THEM: reload page to use the new value ---------- */  
  
  var ZetoSheet = SpreadsheetApp.getActiveSheet()
  var TodayRow = findDate(ZetoSheet, new Date())
  ZetoSheet.getRange(TodayRow, 1).activate()
  
  CheckAndAddDates();
  CheckSum();
}


//handles a user clicking the "Sign Up" button on the spreadsheet or the "Sign Up" menu item
function openDialog() {
  var doc = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1NyxrPYQiOddbxE2IGWc_vvvEUFZctzaaYt1CtMiT0r4/edit#gid=0");
  var html = HtmlService.createHtmlOutputFromFile('UI')
      .setWidth(500)
      .setHeight(445)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Sign Up for a time slot:');
}


function onEdit(e)
{
  var TargetCell = e.range;
  var ZetoSheet = SpreadsheetApp.getActiveSheet();
  //Browser.msgBox(TargetCell.getColumn());
  if (TargetCell.getColumn() == 1 && TargetCell.getRow() == 1)
  {
    Browser.msgBox("Please do not alter the date column as this may affect the scheduler. If you are not the first person to sign up for a day, your date will not be entered. This is expected and is for organizational purposes. Thank you.");
    TargetCell.setValue("Date");
  }
  if (TargetCell.getColumn() == 1 && ZetoSheet.getRange(TargetCell.getRowIndex(), parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == "")
  {
    Browser.msgBox("Please do not alter the date column as this may affect the scheduler. If you are not the first person to sign up for a day, your date will not be entered. This is expected and is for organizational purposes. Thank you.");
    TargetCell.setValue("");
  }
  else if (TargetCell.getColumn() == 1 && ZetoSheet.getRange(TargetCell.getRowIndex(), parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() != "")
  {
    Browser.msgBox("Please do not alter the date column as this may affect the scheduler, thank you.");
    var TargetRow = TargetCell.getRow()-1;
    while (ZetoSheet.getRange(TargetRow, 1).getValue() == "" && TargetRow != 1)
    {
      TargetRow--;
    }
    var newDate = ZetoSheet.getRange(TargetRow, 1).getValue();
    newDate.setDate(newDate.getDate() + 1);
    TargetCell.setValue(newDate);
  }
}
/**/




//this function is called when the Sign Up button is clicked,
//the formObject passed into the function is from the form
//the user filled out in HTML
function SignUpJS(formObject) {
  
  //This variable is the deadline (1-24), plates must be turned in
  //before this hour in order to be run on the desired day. No new
  //signups can occur after the daily deadline
  var deadline = 16; //Current Deadline for plates: 4PM
  //CheckSum();
  
  
  /*-----------------------------------------------
  This block grabs information from the HTML form
  and closes the popup afterwards                */
  var date = formObject.theDate;
  var user = formObject.username;
  var markers = formObject.markers;
  var ladder = formObject.ladder;
  var pType = formObject.ptype;
  var PI = formObject.principleInvestigator;
  var pName = formObject.plateName;
  /*---------------------------------------------*/
  
  
  //Checking for correct minimum user input in form
  if (user == "" || pName == "" || date == "") 
  {
    var html = HtmlService.createHtmlOutputFromFile('Fail-NeedAdditionalInformation')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Sign Up Unsuccessful');
    return;
  }
  
  //acquire current date
  var currentDate = new Date();
  //loading spreadsheet into javascript for use
  var ZetoSheet = SpreadsheetApp.getActiveSheet();
  var FormDate = new Date();
  FormDate = convertSSdate(date);
  
  //check if date selected is in the past, if so cancel sign up
  if (currentDate.getYear() > FormDate.getYear() 
    || (currentDate.getYear() == FormDate.getYear() && currentDate.getMonth() > FormDate.getMonth()) 
    || (currentDate.getYear() == FormDate.getYear() && currentDate.getMonth() == FormDate.getMonth() && currentDate.getDate() > FormDate.getDate()))
    {
      HandlePastSignUp();
      return;
    }
  
  var dateRow = findDate(ZetoSheet, FormDate);
  
  //check to see if user input an invalid date
  if (dateRow == 0)
  {
    HandleInvalidDate();
    return;
  }
  
  
  
  //if the date on the form is today
  if ((currentDate.getDate()) == FormDate.getDate() && currentDate.getMonth() == FormDate.getMonth() && currentDate.getYear() == FormDate.getYear())
  {
    if(currentDate.getHours() < deadline)
      HandleSignUpToday(ZetoSheet, formObject, dateRow);
    else
      HandlePastSignUp();
  }
  
  //else if the date on the form is in the future
  else if (currentDate.getYear() < FormDate.getYear() 
    || (currentDate.getYear() == FormDate.getYear() && currentDate.getMonth() < FormDate.getMonth()) 
    || (currentDate.getYear() == FormDate.getYear() && currentDate.getMonth() == FormDate.getMonth() && currentDate.getDate() < FormDate.getDate()))
  {
    HandleFutureSignUp(ZetoSheet, formObject, dateRow, FormDate);
  }
  
  //else if the date on the form is in the past
  else
  {
    HandlePastSignUp();
  }
  
  return;
}








//this function converts a string of format YYYY-MM-DD into a date javascript class and returns the date
function convertSSdate(formDate) {
  var spl = formDate.split("/");
  var d = new Date(0,0,0,0,0,0,0);
  d.setDate(parseInt(spl[1], 10));
  d.setFullYear(parseInt(spl[2], 10));
  d.setMonth(parseInt(spl[0], 10)-1);
  return d;
}




//this function scans the first column of the spreadsheet looking for the number row in which
//the date given by the HTML form matches the date given by the spreadsheet. It returns the row number
//in integer form of the first row of the date inputted.
function findDate(ZetoSheet, FormDate) {
  var startRow = parseInt(PropertiesService.getScriptProperties().getProperty('STARTROW'), 10)
  
  var referenceColumn = ZetoSheet.getRange(7, 1, ZetoSheet.getLastRow(), 1).getValues()
  var refDate = new Date()
  var lastRow = ZetoSheet.getLastRow()-7
  var i = 0
  
  for (i = 0; i < referenceColumn.length; i++)
  {
    if (referenceColumn[i] == "")
    {
      continue
    }
    refDate = new Date(referenceColumn[i])
    
    if (FormDate.getYear() == refDate.getYear() && FormDate.getMonth() == refDate.getMonth() && FormDate.getDate() == refDate.getDate())
    {

      break
    }
    if (i == lastRow)
      return 0
  }
  return i+7;
}



//this function handles an invalid date selection on the form. This function is 
//called when a date is not found on the spreadsheet and either indicates that
//the date selected was too far in the past, or that more dates need to be 
//added to the spreadsheet
function HandleInvalidDate()
{
   var html = HtmlService.createHtmlOutputFromFile('Fail-DateNotFound')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Invalid Date Selected');
}















/*--------------------------------------------------------------------------------
----------------------This segment handles sign ups on today's date------------*/


//this function will handle a signup on the current day
function HandleSignUpToday(ZetoSheet, formObject, dateRow)
{ 
  //check for if there are no signups on this date yet
  if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == "" || ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 0)
  {
    ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(0);
    //at this point there IS NOT a signup for this date, handle 96 well plate
    if (formObject.ptype == 96)
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(1.5);
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
  
    //at this point there IS NOT a signup for this date, handle 384 well plate
    else if (formObject.ptype = 384)
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(6);
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
  }
  
  
  
  //at this point there IS a signup for this date, handle 96 well plate
  else if (formObject.ptype == 96)
  {
    //check to see if there is enough time to schedule this plate
    if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() > 22.5)
    {
      var html = HtmlService.createHtmlOutputFromFile('Fail-Today-NotEnoughTime')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Sign Up Unsuccessful');
      return false;
    }
    
    //check if initial row is empty from a previous deletion, if so populate it
    if (ZetoSheet.getRange(dateRow, 5).getValue() == "")
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+1.5);
      if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, dateRow);
      }
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
    
    //input proper information into new spreadsheet row and update allocated time
    else
    {
      
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+1.5);
      var oldDateRow = dateRow;
      //this while loop ensures that the order of signup is maintained
      while (ZetoSheet.getRange(dateRow+1, 1).getValue() == "")
        dateRow++;
      ZetoSheet.insertRowAfter(dateRow);
      dateRow++;
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
      if (ZetoSheet.getRange(oldDateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, oldDateRow);
      }
    }
  }
  
  //at this point there IS a signup for this date, handle 384 well plate
  else if (formObject.ptype = 384)
  {
    
    //check to see if there is enough time to schedule this plate
    if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() > 18)
    {
      var html = HtmlService.createHtmlOutputFromFile('Fail-Today-NotEnoughTime')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Sign Up Unsuccessful');
      return false;
    }
    
     //check if initial row is empty from a previous deletion, if so populate it
    if (ZetoSheet.getRange(dateRow, 5).getValue() == "")
    {
      if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, dateRow);
      }
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+6);
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
    
    //input proper information into new spreadsheet row and update allocated time
    else 
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+6);
      var oldDateRow = dateRow;
      //this while loop ensures that the order of signup is maintained
      while (ZetoSheet.getRange(dateRow+1, 1).getValue() == "")
        dateRow++;
      ZetoSheet.insertRowAfter(dateRow);
      dateRow++;
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
      
      if (ZetoSheet.getRange(oldDateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, oldDateRow);
      }
    }
  }
  
  
  var html = HtmlService.createHtmlOutputFromFile('Success-TodayChosen')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Sign Up Successful');
  return true;
}

/*------------------------------------------------------------------------------*/













/*--------------------------------------------------------------------------------
----------------------This segment handles sign ups on a future date------------*/

//this function handles a signup event on a future date.
function HandleFutureSignUp(ZetoSheet, formObject, dateRow, date)
{
  //check for if there are no signups on this date yet
  if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == "" || ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 0)
  {
    ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(0);
    //at this point there IS NOT a signup for this date, handle 96 well plate
    if (formObject.ptype == 96)
    {
      /*ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(1.5);
      string = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      ZetoSheet.getRange(dateRow, 2).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      ZetoSheet.getRange(dateRow, 3).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      ZetoSheet.getRange(dateRow, 4).setValue(string);
      ZetoSheet.getRange(dateRow, 5).setValue(formObject.username);
      ZetoSheet.getRange(dateRow, 6).setValue(formObject.markers);
      ZetoSheet.getRange(dateRow, 7).setValue(formObject.ladder);
      ZetoSheet.getRange(dateRow, 8).setValue(formObject.ptype);
      ZetoSheet.getRange(dateRow, 9).setValue(formObject.principleInvestigator);
      ZetoSheet.getRange(dateRow, 10).setValue(formObject.plateName);
      ZetoSheet.getRange(dateRow, 12).setValue(formObject.email);*/
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(1.5);
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
  
    //at this point there IS NOT a signup for this date, handle 384 well plate
    else if (formObject.ptype = 384)
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(6);
      /*string = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      ZetoSheet.getRange(dateRow, 2).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      ZetoSheet.getRange(dateRow, 3).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      ZetoSheet.getRange(dateRow, 4).setValue(string);
      ZetoSheet.getRange(dateRow, 5).setValue(formObject.username);
      ZetoSheet.getRange(dateRow, 6).setValue(formObject.markers);
      ZetoSheet.getRange(dateRow, 7).setValue(formObject.ladder);
      ZetoSheet.getRange(dateRow, 8).setValue(formObject.ptype);
      ZetoSheet.getRange(dateRow, 9).setValue(formObject.principleInvestigator);
      ZetoSheet.getRange(dateRow, 10).setValue(formObject.plateName);
      ZetoSheet.getRange(dateRow, 12).setValue(formObject.email);*/
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
  }
  
  //at this point there IS a signup for this date, handle 96 well plate
  else if (formObject.ptype == 96)
  {
    //check to see if there is enough time to schedule this plate
    if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() > 22.5)
    {
      var html = HtmlService.createHtmlOutputFromFile('Fail-Today-NotEnoughTime')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Sign Up Unsuccessful');
      return false;
    }
    
    //check if initial row is empty from a previous deletion, if so populate it
    if (ZetoSheet.getRange(dateRow, 5).getValue() == "")
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+1.5);
      if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, dateRow);
      }
      /*string = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      ZetoSheet.getRange(dateRow, 2).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      ZetoSheet.getRange(dateRow, 3).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      ZetoSheet.getRange(dateRow, 4).setValue(string);
      ZetoSheet.getRange(dateRow, 5).setValue(formObject.username);
      ZetoSheet.getRange(dateRow, 6).setValue(formObject.markers);
      ZetoSheet.getRange(dateRow, 7).setValue(formObject.ladder);
      ZetoSheet.getRange(dateRow, 8).setValue(formObject.ptype);
      ZetoSheet.getRange(dateRow, 9).setValue(formObject.principleInvestigator);
      ZetoSheet.getRange(dateRow, 10).setValue(formObject.plateName);
      ZetoSheet.getRange(dateRow, 12).setValue(formObject.email);*/
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
    //input proper information into new spreadsheet row and update allocated time
    else
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+1.5);
      var oldDateRow = dateRow;
      //this while loop ensures that the order of signup is maintained
      while (ZetoSheet.getRange(dateRow+1, 1).getValue() == "")
        dateRow++;
      ZetoSheet.insertRowAfter(dateRow);
      /*string = "=IF(ISBLANK(A"+ dateRow+1 + "), \"\", FLOOR((24 - M" + dateRow+1 + ")/6, 1))";
      ZetoSheet.getRange(dateRow+1, 2).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow+1 + "), \"\", \"or\")";
      ZetoSheet.getRange(dateRow+1, 3).setValue(string);
      string = "=IF(ISBLANK(A"+ dateRow+1 + "), \"\", FLOOR((24 - M" + dateRow+1 + ")/1.5, 1))";
      ZetoSheet.getRange(dateRow+1, 4).setValue(string);
      ZetoSheet.getRange(dateRow+1, 5).setValue(formObject.username);
      ZetoSheet.getRange(dateRow+1, 6).setValue(formObject.markers);
      ZetoSheet.getRange(dateRow+1, 7).setValue(formObject.ladder);
      ZetoSheet.getRange(dateRow+1, 8).setValue(formObject.ptype);
      ZetoSheet.getRange(dateRow+1, 9).setValue(formObject.principleInvestigator);
      ZetoSheet.getRange(dateRow+1, 10).setValue(formObject.plateName);
      ZetoSheet.getRange(dateRow+1, 12).setValue(formObject.email);*/
      dateRow++;
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
      
      if (ZetoSheet.getRange(oldDateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, oldDateRow);
      }
    }
  }
  
  //at this point there IS a signup for this date, handle 384 well plate
  else if (formObject.ptype = 384)
  {
    
    //check to see if there is enough time to schedule this plate
    if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() > 18)
    {
      var html = HtmlService.createHtmlOutputFromFile('Fail-Today-NotEnoughTime')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Sign Up Unsuccessful');
      return false;
    }
    
    //check if initial row is empty from a previous deletion, if so populate it
    if (ZetoSheet.getRange(dateRow, 5).getValue() == "")
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+6);
      if (ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, dateRow);
      }
      /*ZetoSheet.getRange(dateRow, 5).setValue(formObject.username);
      ZetoSheet.getRange(dateRow, 6).setValue(formObject.markers);
      ZetoSheet.getRange(dateRow, 7).setValue(formObject.ladder);
      ZetoSheet.getRange(dateRow, 8).setValue(formObject.ptype);
      ZetoSheet.getRange(dateRow, 9).setValue(formObject.principleInvestigator);
      ZetoSheet.getRange(dateRow, 10).setValue(formObject.plateName);
      ZetoSheet.getRange(dateRow, 12).setValue(formObject.email);*/
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
    }
    
    //input proper information into new spreadsheet row and update allocated time
    else
    {
      ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(dateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue()+6);
      var oldDateRow = dateRow;
      //this while loop ensures that the order of signup is maintained
      while (ZetoSheet.getRange(dateRow+1, 1).getValue() == "")
        dateRow++;
      ZetoSheet.insertRowAfter(dateRow);
      /*ZetoSheet.getRange(dateRow+1, 5).setValue(formObject.username);
      ZetoSheet.getRange(dateRow+1, 6).setValue(formObject.markers);
      ZetoSheet.getRange(dateRow+1, 7).setValue(formObject.ladder);
      ZetoSheet.getRange(dateRow+1, 8).setValue(formObject.ptype);
      ZetoSheet.getRange(dateRow+1, 9).setValue(formObject.principleInvestigator);
      ZetoSheet.getRange(dateRow+1, 10).setValue(formObject.plateName);
      ZetoSheet.getRange(dateRow+1, 12).setValue(formObject.email);*/
      dateRow++;
      string1 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ dateRow + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ dateRow + "), \"\", FLOOR((24 - M" + dateRow + ")/1.5, 1))";
      values = [
        [string1, string2, string3, formObject.username, formObject.markers, formObject.ladder, formObject.ptype, formObject.principleInvestigator, formObject.plateName, "", formObject.email]
        ];
      ZetoSheet.getRange(dateRow, 2, 1, 11).setValues(values);
      if (ZetoSheet.getRange(oldDateRow, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() == 24)
      {
        DemarcateFullDate(ZetoSheet, oldDateRow);
      }
    }
  }
  
  var html = HtmlService.createHtmlOutputFromFile('Success-FutureDateChosen')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Sign Up Successful for ' + (date.getMonth()+1) + "/" + date.getDate() + "/" + date.getYear());
  
  return true;
}
/*------------------------------------------------------------------------------*/
















/*--------------------------------------------------------------------------------------
-This segment handles sign ups on a passed date, this includes the same day after 4PM-*/


//this function handles a signup event on a past date.
function HandlePastSignUp()
{
  var html = HtmlService.createHtmlOutputFromFile('Fail-PastDateChosen')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Sign Up Unsuccessful');
  return false;
}

/*------------------------------------------------------------------------------*/

















/*--------------------------------------------------------------------------------------
-----------This segment handles someone clicking the "Remove Plate" button--------------
----------  RemovePlateButton() ----> RemovePlate() ----> PlateRemoved() ------------ */

//this function will prompt the clicker to enter an "admin password" to remove a plate
//they have added to the scheduler
function RemovePlateButton()
{
  var html = HtmlService.createHtmlOutputFromFile('RemovePlateButton')
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Remove Plate From Scheduler');
}

//we are making password protection not necessary to remove plates, to reimplement remove comments from this section and call "RemovePlateButton" from drop down menu option
function RemovePlate(/*formObject*/)
{
  //change me to change the password!
    var password = "wrsggl";
  //^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  
 /* if (formObject.pwd != password)
  {
    var html = HtmlService.createHtmlOutput('Please retry or contact Johnson Hall 291B to remove this plate').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Password Incorrect!');
  }
  else
  { */
  var html = HtmlService.createHtmlOutput('To ensure that the wrong plate is not removed, please input the row number, username, and plate name of the plate to be removed.<br><br> <form>Row number of the plate being removed: <input type="text" name="rowNum"><hr> User name of plate being removed: <input type="text" name="username"><hr> Name of plate being removed: <input type="text" name="plateName"><br> <input type="button" value="OK" onclick="google.script.run.PlateRemoved(this.parentNode).host.close()" /> </form>').setWidth(500).setHeight(200).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Remove a Plate');
  //}
}

//this function, having acquired the row to remove, removes the given row
function PlateRemoved(formObject)
{
  var rowNumber = parseInt(formObject.rowNum, 10);
  if (rowNumber <= 6)
  {
    Browser.msgBox("You cannot remove the header rows!");
    return;
  }
  var ZetoSheet = SpreadsheetApp.getActiveSheet();
  if (ZetoSheet.getRange(rowNumber, 5).getValue() != formObject.username || ZetoSheet.getRange(rowNumber, 10).getValue() != formObject.plateName)
  {
    Browser.msgBox("The plate in row "+rowNumber+" cannot be removed because either the user name or plate name given does not match the plate that is in that row. It is possible that a user added their plate while you were deleting yours, so please retry removing your plate and make sure you enter the proper values for user name and plate name. Thank you!");
    return;
  }
  if (ZetoSheet.getRange(rowNumber, 1).getValue() == "")
  {
    MailApp.sendEmail('wrsggl@gmail.com', (ZetoSheet.getRange(rowNumber, 5).getValue()+"'s Plate Deleted"), (ZetoSheet.getRange(rowNumber, 5).getValue()+"'s plate was deleted on: "+Date()+"\n"+"Row Number: "+rowNumber+"\n"+"Data: "+ZetoSheet.getRange(rowNumber, 1).getValue()+" "+ZetoSheet.getRange(rowNumber, 5).getValue()+" "+ZetoSheet.getRange(rowNumber, 6).getValue()+" "+ZetoSheet.getRange(rowNumber, 7).getValue()+" "+ZetoSheet.getRange(rowNumber, 8).getValue()+" "+ZetoSheet.getRange(rowNumber, 9).getValue()+" "+ZetoSheet.getRange(rowNumber, 10).getValue()+" "+ZetoSheet.getRange(rowNumber, 11).getValue()+" "+ZetoSheet.getRange(rowNumber, 12).getValue()) )
    var oldRow = rowNumber;
    var ptype = ZetoSheet.getRange(rowNumber, 8).getValue();
    while(ZetoSheet.getRange(rowNumber, 1).getValue() == "" && rowNumber != 0)
    {
      rowNumber--;
    }
    if(rowNumber == 0)
      return;
    var TotalHours = ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue();
    if(ptype=="96")
    {
      ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() - 1.5);
    }
    else if (ptype == "384")
    {
      ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() - 6);
    }
    ZetoSheet.deleteRow(oldRow);
    if (TotalHours == 24)
    {
      DateNoLongerFull(ZetoSheet, rowNumber);
    }
  }
  else
  {
    MailApp.sendEmail('wrsggl@gmail.com', (ZetoSheet.getRange(rowNumber, 5).getValue()+"'s Plate Deleted"), (ZetoSheet.getRange(rowNumber, 5).getValue()+"'s plate was deleted on: "+Date()+"\n"+"Row Number: "+rowNumber+"\n"+"Data: "+ZetoSheet.getRange(rowNumber, 1).getValue()+" "+ZetoSheet.getRange(rowNumber, 5).getValue()+" "+ZetoSheet.getRange(rowNumber, 6).getValue()+" "+ZetoSheet.getRange(rowNumber, 7).getValue()+" "+ZetoSheet.getRange(rowNumber, 8).getValue()+" "+ZetoSheet.getRange(rowNumber, 9).getValue()+" "+ZetoSheet.getRange(rowNumber, 10).getValue()+" "+ZetoSheet.getRange(rowNumber, 11).getValue()+" "+ZetoSheet.getRange(rowNumber, 12).getValue()) )
    var TotalHours = ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue();
    if (ZetoSheet.getRange(rowNumber, 8).getValue() == "96")
      ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() - 1.5);
    else if (ZetoSheet.getRange(rowNumber, 8).getValue() == "384")
      ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setValue(ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).getValue() - 6);
    
    //if this plate is the only one being run today, reset all values to default
    if (ZetoSheet.getRange(rowNumber+1, 1).getValue() != "" || rowNumber == ZetoSheet.getLastRow())
    {
      /*
      string = "=IF(ISBLANK(A"+ rowNumber + "), \"\", FLOOR((24 - M" + rowNumber + ")/6, 1))";
      ZetoSheet.getRange(rowNumber, 2).setValue(string);
      string = "=IF(ISBLANK(A"+ rowNumber + "), \"\", \"or\")";
      ZetoSheet.getRange(rowNumber, 3).setValue(string);
      string = "=IF(ISBLANK(A"+ rowNumber + "), \"\", FLOOR((24 - M" + rowNumber + ")/1.5, 1))";
      ZetoSheet.getRange(rowNumber, 4).setValue(string);
      ZetoSheet.getRange(rowNumber, 5).setValue("");
      ZetoSheet.getRange(rowNumber, 6).setValue("");
      ZetoSheet.getRange(rowNumber, 7).setValue("");
      ZetoSheet.getRange(rowNumber, 8).setValue("");
      ZetoSheet.getRange(rowNumber, 9).setValue("");
      ZetoSheet.getRange(rowNumber, 10).setValue("");
      ZetoSheet.getRange(rowNumber, 11).setValue("");
      ZetoSheet.getRange(rowNumber, 12).setValue("");
      */
      editRange = ZetoSheet.getRange(rowNumber, 2, 1, 11);
      string1 = "=IF(ISBLANK(A"+ rowNumber + "), \"\", FLOOR((24 - M" + rowNumber + ")/6, 1))";
      string2 = "=IF(ISBLANK(A"+ rowNumber + "), \"\", \"or\")";
      string3 = "=IF(ISBLANK(A"+ rowNumber + "), \"\", FLOOR((24 - M" + rowNumber + ")/1.5, 1))";
      values = [
        [string1, string2, string3, "", "", "", "", "", "", "", ""]
        ];
      editRange.setValues(values);
    }
    //otherwise...
    else if(ZetoSheet.getRange(rowNumber+1, 1).getValue() == "" && rowNumber+1 != ZetoSheet.getLastRow())
    {
      ZetoSheet.getRange(rowNumber+1, 1).setValue(ZetoSheet.getRange(rowNumber, 1).getValue());
      ZetoSheet.getRange(rowNumber+1, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'))).setValue(ZetoSheet.getRange(rowNumber, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'))).getValue());
      ZetoSheet.deleteRow(rowNumber);
      string = "=IF(ISBLANK(A"+ rowNumber + "), \"\", FLOOR((24 - M" + rowNumber + ")/6, 1))";
      ZetoSheet.getRange(rowNumber, 2).setValue(string);
      string = "=IF(ISBLANK(A"+ rowNumber + "), \"\", \"or\")";
      ZetoSheet.getRange(rowNumber, 3).setValue(string);
      string = "=IF(ISBLANK(A"+ rowNumber + "), \"\", FLOOR((24 - M" + rowNumber + ")/1.5, 1))";
      ZetoSheet.getRange(rowNumber, 4).setValue(string);

    }
    if (TotalHours == 24)
    {
      DateNoLongerFull(ZetoSheet, rowNumber);
    }
  }
  //CheckSum();
  var html = HtmlService.createHtmlOutput("Plate Removal Successful!<br> <input type=\"button\" value=\"Close\" onclick=\"google.script.host.close()\"  />").setWidth(500).setHeight(200).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Done');
}
         
/*---------------------------------------------------------------------------------*/








/*------------------------------------------------------------------------------------------------------------------
------------------This segment handles someone clicking the "Export and Remove Data" menu Option--------------------
---------- ExportAndRemoveMenuOption() ----> ExportAndRemoveHelper1() ----> ExportAndRemoveHelper2() ------------ */
function ExportAndRemoveMenuOption()
{
  var html = HtmlService.createHtmlOutput('Be aware that the ability to export and remove data is an admin function, and will require an admin password to continue: <br> <form> Password <input type="password" name="pwd"> <br> <input type="button" value="Continue" onclick="google.script .run.ExportAndRemoveHelper1(this.parentNode) .host.close()" /> <input type="button" value="Cancel" onclick="google.script.host.close()" /> </form>').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Export and Remove Data');
}

function ExportAndRemoveHelper1(formObject)
{
  var password = PropertiesService.getScriptProperties().getProperty('PASSWORD');
   if (formObject.pwd != password)
  {
    var html = HtmlService.createHtmlOutput('Please retry or contact Johnson Hall 291B to export this data').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Password Incorrect!');
  }
  else
  {
    var html = HtmlService.createHtmlOutput('Please be aware that all spreadsheet data up to but not including the current date will be exported to the Zeto Sheet Archive and be removed from this spreadsheet. <input type="button" value="OK" onclick="google.script .run.ExportAndRemoveHelper2() .host.close()" /> </form>').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Password correct!');
  }
}

function ExportAndRemoveHelper2()
{
  var currentDate = new Date();
  var FormDate = new Date();
  //this sets the removal date to the day before today's date
  FormDate.setDate(FormDate.getDate() - 1);
  //check to make sure date given to export is in the past
  if (currentDate.getYear() > FormDate.getYear() 
    || (currentDate.getYear() == FormDate.getYear() && currentDate.getMonth() > FormDate.getMonth()) 
    || (currentDate.getYear() == FormDate.getYear() && currentDate.getMonth() == FormDate.getMonth() && currentDate.getDate() > FormDate.getDate()))
    {
      var ZetoSheet = SpreadsheetApp.getActiveSheet();
      var dateRow = findDate(ZetoSheet, FormDate);
      //if the findDate function cannot find the given date, it returns zero, this is the handler for that occurance.
      if (dateRow == 0)
      {
        var html = HtmlService.createHtmlOutput('The date you have selected was not found on the spreadsheet, please check that the date is present and try again.')
          .setWidth(500)
          .setHeight(500)
          .setSandboxMode(HtmlService.SandboxMode.IFRAME);
        SpreadsheetApp.getUi().showModalDialog(html, 'Invalid Date');
        return;
      }
      //get to the proper row including any entries below the date row
      while(ZetoSheet.getRange(dateRow+1, 1).getValue() == "")
      {
        dateRow++;
      }
      
      //open the Archive spreadsheet
      var Archive = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SAVEID'));
      //get the specific sheet that has all the data
      var ArchiveSheet = Archive.getSheetByName("Archive");
      
      var ArchiveRow = 1;
      //get to the first empty row in the archive to begin recording
      while (ArchiveSheet.getRange(ArchiveRow, 5).getValue() != "")
      {
        ArchiveRow++;
      }
      var currentRow = 7;
      var currentCol = 1;
      //save this value in a variable to avoid multiple calls to the properties service
      var timeCol = parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10);

      while(currentRow <= dateRow)
      {
        while (currentCol <= timeCol)
        {
          ArchiveSheet.getRange(ArchiveRow, currentCol).setValue(ZetoSheet.getRange(currentRow, currentCol).getValue());
          currentCol++;
        }
        currentRow++;
        ArchiveRow++;
        currentCol=1;
      }

      ZetoSheet.deleteRows(7, dateRow-6)
      
      var html = HtmlService.createHtmlOutput('The file has been exported to the Zeto Scheduler Archive.')
        .setWidth(500)
        .setHeight(500)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi().showModalDialog(html, 'Export Successful');
    }
  //handle current or future date
  else
  {
    var html = HtmlService.createHtmlOutput('The date you have selected is either today\'s date, or a date in the future. Please select a date in the past to export and remove.').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Invalid Date');
    return;
  }
    
}

/*----------------------------------------------------------------------------------------------------------------*/
/* This segment deals with coloring/decoloring days that have the maximum amount of plates */


function DemarcateFullDate(ZetoSheet, dateRow)
{
  var endRow = dateRow;
  while (ZetoSheet.getRange(endRow+1, 1).getValue() == "" && endRow != ZetoSheet.getLastRow())
  {
    endRow++;
  }
  ZetoSheet.getRange(dateRow, 1, endRow-dateRow+1, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setBackgroundRGB(200, 50, 50);
}

function DateNoLongerFull(ZetoSheet, rowNumber)
{
  var endRow = rowNumber;
  while (ZetoSheet.getRange(endRow+1, 1).getValue() == "" && endRow != ZetoSheet.getLastRow())
  {
    endRow++;
  }
  ZetoSheet.getRange(rowNumber, 1, endRow-rowNumber+1, parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10)).setBackgroundRGB(255, 255, 255)
}
/*-----------------------------------------------------------------------------------------------------------------*/



function CheckSum()
{
  var ZetoSheet = SpreadsheetApp.getActiveSheet();
  var currentRow = parseInt(PropertiesService.getScriptProperties().getProperty('STARTROW'), 10);
  var dateRow = currentRow;
  var checkSum = 0;
  var timeCol = parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10);
  var ptypeRow = parseInt(PropertiesService.getScriptProperties().getProperty('PLATETYPEROW'), 10);
  //Browser.msgBox("Start checksum");
  while (currentRow != ZetoSheet.getLastRow())
  {
    checkSum = 0;
    dateRow = currentRow;
    if (ZetoSheet.getRange(currentRow, ptypeRow).getValue() == 384)
      checkSum = 6;
    else if (ZetoSheet.getRange(currentRow, ptypeRow).getValue() == 96)
      checkSum = 1.5;
    
    currentRow++;
    while (ZetoSheet.getRange(currentRow, 1).getValue() == "")
    {
      if (ZetoSheet.getRange(currentRow, 1, 1, timeCol).isBlank())
      {
        //Browser.msgBox("Blank row found at row "+currentRow+". Removing the row now.");
        ZetoSheet.deleteRow(currentRow);
        continue;
      }
      if (ZetoSheet.getRange(currentRow, ptypeRow).getValue() == 384)
        checkSum += 6;
      else if (ZetoSheet.getRange(currentRow, ptypeRow).getValue() == 96)
        checkSum += 1.5;
      currentRow++;
    }
    if (checkSum != ZetoSheet.getRange(dateRow, timeCol).getValue())
    {
      //Browser.msgBox("Row "+dateRow+" has a discrepancy.\nOld value: "+ZetoSheet.getRange(dateRow, timeCol).getValue()+"\nNew value: "+checkSum);
      
      //decolor/color the row if it is no longer full/now is full
      if(ZetoSheet.getRange(dateRow, timeCol).getValue() == 24)
        DateNoLongerFull(ZetoSheet, dateRow);
      else if(checkSum == 24)
        DemarcateFullDate(ZetoSheet, dateRow);
      
      ZetoSheet.getRange(dateRow, timeCol).setValue(checkSum);
    }
  }
  //Browser.msgBox("End checksum");
}

function CheckAndAddDates()
{
  var ZetoSheet = SpreadsheetApp.getActiveSheet();
  var timeCol = parseInt(PropertiesService.getScriptProperties().getProperty('TIMECOLUMN'), 10);
  var currentRow = 7;
  while (!(ZetoSheet.getRange(currentRow+1, 1, 1, timeCol-1).isBlank() && ZetoSheet.getRange(currentRow+2, 1, 1, timeCol-1).isBlank() && ZetoSheet.getRange(currentRow+3, 1, 1, timeCol-1).isBlank()))
  {
    currentRow++;
  }
  var lastDate = ZetoSheet.getRange(currentRow, 1).getValue();
  var todayDate=  new Date()
  var thresholdDate = todayDate;
  for (var i = 0 ; i < 62 ; i++)
  {
    thresholdDate.setDate(thresholdDate.getDate()+1);
  }
  var todayDate=  new Date();
  if (lastDate < thresholdDate)
  {
    while (lastDate < thresholdDate)
    {
      currentRow++;
      lastDate.setDate(lastDate.getDate()+1);
      ZetoSheet.getRange(currentRow, 1).setValue(lastDate);
      ZetoSheet.getRange(currentRow, 2).setValue("=IF(ISBLANK(A"+currentRow+"), \"\", FLOOR((24 - M"+currentRow+")/6.0, 1))");
      ZetoSheet.getRange(currentRow, 3).setValue("=IF(ISBLANK(A"+currentRow+"), \"\", \"or\")");
      ZetoSheet.getRange(currentRow, 4).setValue("=IF(ISBLANK(A"+currentRow+"), \"\", FLOOR((24 - M"+currentRow+")/1.5, 1))");
      ZetoSheet.getRange(currentRow, timeCol).setValue(0);
      
    }
  }
}

/****************************************************************
************ Admin -> Toggle Zeto Access Implementation *********
** Uses a value in the spreadsheet to change the login message */
function ToggleZetoAccess()
{
  var html = HtmlService.createHtmlOutput('Be aware that the ability to toggle Zeto access is an admin function, and will require an admin password to continue: <br> <form> Password <input type="password" name="pwd"> <br> <input type="button" value="Continue" onclick="google.script .run.ToggleHelper(this.parentNode) .host.close()" /> <input type="button" value="Cancel" onclick="google.script.host.close()" /> </form>').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Toggle Zeto Access');
}

//Helper function called by the html above
function ToggleHelper(formObject)
{
  var password = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (formObject.pwd != password)
  {
    var html = HtmlService.createHtmlOutput('Please retry or contact Johnson Hall 291B to toggle Zeto access').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Password Incorrect!');
  }
  else
  {
    var ZetoSheet = SpreadsheetApp.getActiveSheet();
    var Arow = PropertiesService.getScriptProperties().getProperty('ACCESSTOGGLEROW');
    var Acol = PropertiesService.getScriptProperties().getProperty('ACCESSTOGGLECOL');
    //Browser.msgBox(Arow);
    var val = ZetoSheet.getRange(Arow, Acol).getValue();
    val += 1;
    val %= 2;
    ZetoSheet.getRange(Arow, Acol).setValue(val);
    var html = HtmlService.createHtmlOutput('Zeto Access successfully toggled. <input type="button" value="OK" onclick="google.script.host.close()" /> </form>').setWidth(500).setHeight(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Password correct!');
  }
}

function TestFunction()
{
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  Browser.msgBox('Tomorrow is: '+ tomorrow);
}


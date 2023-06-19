/**
 * This function handles the onChange trigger event. It checks for the user creating a new sheet by selecting from the Menu 
 * bar File -> Import and then choosing the option to "Insert New Sheet" to the spreadsheet. The data is then processed and
 * prepared in a format ready for import into Adagio. This function also checks if new rows are added to the export sheet 
 * and ensures that they start with the comment line marker 'C' for valid input into Adagio.
 * 
 * @param {Event} e : The event object 
 */
function onChange(e)
{
  if (e.changeType === 'INSERT_GRID') // A new sheet has been created
    processImportedData(e)
  else if (e.changeType === 'INSERT_ROW') // A row has been added
    makeCommentLines(e)
}

/**
 * This function handles the onEdit trigger event. It checks for someone performing a search on one of the search pages.
 * It also updates the PNT descriptions on the vendor pages if a user is making a change on a particlar vendor's sheet.
 * This function also manages the edits that occur on the Export sheets, for instance a user may select File -> Import
 * and then choose to "Append Rows to Current Sheet", in which case the sheet is reformated and the data is processed and
 * prepared in a format ready for import into Adagio. The user may also edit SKUs from the sheet or change the vendor.
 * 
 * @param {Event} e : The event object 
 */
function onEdit_Installed(e)
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getSheetName()

  if (sheetName === 'Search 1' || sheetName === 'Search 2')
    search(e, sheet)
  else if (sheetName === 'Export 1') 
    manageExportSheetEdit(e, 1)
  else if (sheetName === 'Export 2')
    manageExportSheetEdit(e, 2)
  else if (sheetName === 'Grundens' || sheetName === 'Helly Hansen' || sheetName === 'Xtratuf' || sheetName === 'Yeti') // Update
    updatePntDescription(e, sheet, sheetName);
}

/**
 * This function is run when an html web app is launched. In our case, when the modal dialog box is produced at 
 * the point a user has clicked the done button inorder to produce the csv file.
 * 
 * @param {Event} e : The event object 
 */
function doGet(e)
{
  if (e.parameter['sheetName'] !== undefined) // The request parameter
  {
    const sheetName = e.parameter['sheetName'];

    if (sheetName === 'Last Export 1' || sheetName === 'Last Export 2')
      return downloadCsvFile(sheetName)
  }

  return HtmlService.createHtmlOutputFromFile('SuccessfulDownload');
}

/**
 * This function places a custom menu at the top of the page which allows the user to run the addVendor, refresh, and undone functions.
 */
function onOpen()
{
  var menuEntries = [ {name: "Add / Update Vendor",           functionName: "addVendor"}, null, 
                      {name: "Refresh (Update Missing SKUs)", functionName: "refresh"},   null, 
                      {name: "Undone",                        functionName: "undone"},    null,
                      {name: "Update Inventory",              functionName: "importInventory"},
                      {name: "Update Yeti UPCs",              functionName: "importYetiUPCs"}];
  SpreadsheetApp.getActive().addMenu("Pacific Net & Twine Custom Menu Options", menuEntries);
}

/**
 * This function moves data from the current search sheet to the chosen export sheet by identifying which items have quantities typed next to them. 
 * 
 * @param {String} sheetName : The name of the sheet
 * @author Jarren Ralf
 */
function addToExport(sheetName)
{
  const spreadsheet = SpreadsheetApp.getActive();
  const exportSheet = spreadsheet.getSheetByName(sheetName);
  const searchSheet = spreadsheet.getActiveSheet();
  const finalDataRow = searchSheet.getLastRow();
  const col = 2;    // The column that the values on the search page starts on
  const numCols = 2 // The number of columns of the values
  const numHeaders = 4;
  const numItems = finalDataRow - numHeaders;
  const items = searchSheet.getSheetValues(numHeaders + 1, col, numItems, numCols);
  const QTY = 1, DESCRIPTION = 0; // The array indecies 
  var itemToOrder;

  const exportData = items.filter(item => {
    itemToOrder = item[1] !== ''; // If the item has an "Order Quantity" then we add it to the PO
    if (itemToOrder)
    {
      item.push(item[QTY], item[DESCRIPTION]) // Add two columns for the quantity and the description
      item[1] = item[DESCRIPTION].split(' - ', 1)[0] // SKU number
      item[0] = 'R'; // Adagio 'Receiving' line marker
    }
    return itemToOrder
  })

  if (exportData.length !== 0)
  {
    exportSheet.getRange(exportSheet.getLastRow() + 1, 1, exportData.length, exportData[0].length).activate().setNumberFormat('@').setHorizontalAlignment('left')
      .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
      .setValues(exportData)
    searchSheet.getRange(5, 3, numItems).clearContent(); // Clear the quantities
  }
  else
  {
    var activeRanges = searchSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
    var firstRows = [], lastRows = [], numRows = [], itemValues = [[[]]];

    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < activeRanges.length; r++)
    {
      firstRows[r] = activeRanges[r].getRow();
      lastRows[r] = activeRanges[r].getLastRow()
    }

    var     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
    var lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges

    if (row > numHeaders && lastRow <= finalDataRow) // If the user has not selected an item, alert them with an error message
    {   
      for (var r = 0; r < activeRanges.length; r++)
      {
           numRows[r] = lastRows[r] - firstRows[r] + 1;
        itemValues[r] = searchSheet.getSheetValues(firstRows[r], col, numRows[r], numCols);
      }

      // Concatenate all of the item values as a 2-D array, strip the skus off and add the row markers
      var itemVals = [].concat.apply([], itemValues).map(item => {
        item.push(null, item[DESCRIPTION]) // Add two columns for the quantity and the description
        item[1] = item[DESCRIPTION].split(' - ', 1)[0] // SKU number
        item[0] = 'R'; // Adagio 'Receiving' line marker
        return item;
      }); 

      exportSheet.getRange(exportSheet.getLastRow() + 1, 1, itemVals.length, itemVals[0].length).activate().setNumberFormat('@').setHorizontalAlignment('left')
        .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
        .setValues(itemVals); // Move the item values to the destination sheet

      spreadsheet.toast('Your current selection was moved to the export page because you didn\'t add any quantities.', 'Missing Quantities', 10)
    }
    else
      Browser.msgBox('Please add some order quantities or select the data you want to export.')
  }
}

/**
 * This function moves items to the Export 1 sheet.
 * 
 * @author Jarren Ralf
 */
function addToExport1()
{
  addToExport('Export 1');
}

/**
 * This function moves items to the Export 2 sheet.
 * 
 * @author Jarren Ralf
 */
function addToExport2()
{
  addToExport('Export 2');
}

/**
 * This function adds or updates vendor information on the hidden vendor sheet. If the user runs the function on one of the export sheets,
 * then the function checks if the vendor number and name cells have been edited in order to add that information to the list or update current info.
 * If the user runs the function from another sheet, then a ui prompt is launch and the user types in the appropriate info.
 * 
 * @author Jarren Ralf
 */
function addVendor()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const vendorSheet = spreadsheet.getSheetByName('Vendors');
  const vendorsRng = vendorSheet.getRange(2, 1, vendorSheet.getLastRow() - 1, 2);
  const vendors = vendorsRng.getValues();
  const sheetName = sheet.getSheetName();
  const ui = SpreadsheetApp.getUi();
  const blue = '#e0e9f9'

  if (sheetName === 'Export 1' || sheetName === 'Export 2') // If the user runs the function while on the Export page, then check cells B1 and D1, venndor number ad name etc,s
  {
    var vendorInfoRng = sheet.getRange(1, 3, 1, 3);
    var vendorInfo = vendorInfoRng.getValues()[0]; // Vendor number and name
    vendorInfo[2] = toProper(vendorInfo[2]); // Make the vendor proper case
    vendorInfoRng.setHorizontalAlignments([['center', 'left', 'center']]).setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSizes([[13, 13, 18]]).setFontStyle('normal')
      .setFontWeight('bold').setVerticalAlignment('middle').setBackgrounds(blue).setValues([vendorInfo])

    updateVendor(vendorInfo[2], vendorInfo[0], vendors, vendorsRng, vendorSheet, spreadsheet, ui)
  }
  else // If not on the Export page, then a prompt will be displayed
  {
    const response = ui.prompt('New Vendor', 'Please type the new vendor name followed by 2 spaces and the vendor number.', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK)
    {
      var vendorInfo = response.getResponseText().split('  ');

      if (vendorInfo.length !== 2)
        ui.alert('Invalid Input', 'Please try and input the vendor again by typing the new vendor name followed by 2 spaces and the vendor number.', ui.ButtonSet.OK)
      else
      {
        vendorInfo[0] = toProper(vendorInfo[0]);
        updateVendor(vendorInfo[0], vendorInfo[1], vendors, vendorsRng, vendorSheet, spreadsheet, ui)
      }
    }
  }
}

/**
 * This function creates the trigger that will import the inventory every day.
 * 
 * @author Jarren Ralf
 */
function createTriggers()
{
  ScriptApp.newTrigger('importInventory').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('importYetiUPCs').timeBased().atHour(9).everyDays(1).create(); 
  ScriptApp.newTrigger('updateVendors').timeBased().atHour(10).everyDays(1).create();
  ScriptApp.newTrigger('onChange').forSpreadsheet('1WB8DU1rAoRLr3a9t7K5Aa3wCOl08-pqgeV_7Yk4ink8').onChange().create()
  ScriptApp.newTrigger('onEdit_Installed').forSpreadsheet('1WB8DU1rAoRLr3a9t7K5Aa3wCOl08-pqgeV_7Yk4ink8').onEdit().create()
}

/**
 * This function clears the current export sheet of the current data and resets the purchase order header.
 * 
 * @param  {Object[][]}        data : The export data.
 * @param  {Range[][]}        range : The range on the export page that the export data came from
 * @param    {Sheet}          sheet : The name of the sheet.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function clearExportSheet(data, range, sheet, spreadsheet)
{
  const newValues = [...Array(data.length)].map(e => new Array(data[0].length).fill('')); // This array of blanks is used as if to clear the contents of the sheet
  newValues[0][0] = 'H'
  newValues[1][0] = 'C'
  newValues[0][3] = 'Type your order description in this cell. (40 characters max)'
  newValues[1][1] = 'Type your order comments in this cell. (75 characters max)'

  data = data.map(v => {if (v[0] == '') v[0] = 'R'; return v;}) // Make sure all of rows start with R
  data[0][0] = 'H'; // Header
  data[1][0] = 'C'; // Comment Line

  // Remove the first comment line if the user didn't use it
  if (data[1][1] === '' || data[1][1] === 'Type your order comments in this cell. (75 characters max)')
    data.splice(1, 1);

  spreadsheet.getSheetByName('Last Export ' + sheet.getSheetName().split(' ', 2)[1]).clearContents().getRange(1, 1, data.length, data[0].length)
    .setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal')
    .setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(data)

  range.setNumberFormat('@').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle')
    .setValues(newValues)
  sheet.getRange('E2').activate(); // Take the user to user selection dropdown
}

/**
 * This function clears the quantities column on the search page.
 * 
 * @author Jarren Ralf
 */
function clearSearchSheetQTYs()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getLastRow() - 4 !== 0) sheet.getRange(5, 3, sheet.getLastRow() - 4).clearContent()
}

/**
 * This function deletes all of the Triggers associated with the project.
 * 
 * @author Jarren Ralf
 */
function deleteAllTriggers()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger))
}

/**
 * This function downloads the current export sheet as a csv file in order for quick import into the Adagio purchase order system.
 * It checks if there is any invalid information in the data that is about to be download and prompts the user to rectify the issues.
 * It will also clear the page and reset all of the header data, including the user marker.
 * 
 * @author Jarren Ralf
 */
function done()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  const values = range.getValues();
  const backgroundColours = range.getBackgrounds();
  const ok_Cancel_ButtonSet = ui.ButtonSet.OK_CANCEL, cancel = ui.Button.CANCEL;
  const RED = '#f4c7c3', YELLOW = '#fce8b2';
  var response, illegalSkus = [], numAlerts = 0, sideBarMessages = [];
  
  loop: for (var i = 0; i < backgroundColours.length; i++) // Loop throw the rows
  {
    for (var j = 1; j < backgroundColours[0].length; j++) // Loop through the columns
    {
      if (backgroundColours[i][j] === RED) // If cell colour is red
      {
        numAlerts++; // Counting the number of alerts is used to check when an export is completed with no alert messages needed

        if (i === 0) // Check the header row
        {
          if (j === 1) // PO number
          {
            var col = 2;

            if (values[i][j].toString().includes('-'))
              response = ui.alert('⚠️ Change Required ⚠️', 'Please remove the dash (-), from the PO number.\n\nThe correct format is \'PO0#####.\'\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else if (values[i][j].toString().includes(' '))
              response = ui.alert('⚠️ Change Required ⚠️', 'Please remove the spaces, from the PO number.\n\nThe correct format is \'PO0#####.\'\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else if (values[i][j].toString().length < 8)
              response = ui.alert('⚠️ Change Required ⚠️', 'The PO number does not have enough characters.\n\nThe correct format is \'PO0#####.\'\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else if (values[i][j].toString().length > 8)
              response = ui.alert('⚠️ Change Required ⚠️', 'The PO number has too many characters.\n\nThe correct format is \'PO0#####.\'\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else if (isNaN(values[i][j].toString().substring(3))) // The last 5 digits of the PO are non-numeric
              response = ui.alert('⚠️ Change Required ⚠️', 'The PO number contains non-numeric values.\n\nThe correct format is \'PO0#####.\'\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else
              response = ui.alert('⚠️ Change Required ⚠️', 'The PO number appears to be invalid.\n\nThe correct format is \'PO0#####.\'\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)

            break loop;
          }
          else if (j === 2) // Vendor number
          {
            var col = 3;

            if (values[i][j].toString() === '')
              response = ui.alert('⚠️ Change Required ⚠️', 'Adagio requires you to enter a valid vendor number.\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else if (values[i][j].toString().length < 6)
              response = ui.alert('⚠️ Change Required ⚠️', 'The vendor number does not have enough characters (6 numerals only).\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else if (values[i][j].toString().length > 6)
              response = ui.alert('⚠️ Change Required ⚠️', 'The vendor number has too many characters (6 numerals only).\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else if (isNaN(values[i][j].toString())) // The vendor number contains non-numeric values
              response = ui.alert('⚠️ Change Required ⚠️', 'The vendor number must only contain numerical values.\n\n' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)
            else
              response = ui.alert('⚠️ Change Required ⚠️', 'The vendor number appears to be invalid.' + 
                'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
                ok_Cancel_ButtonSet)

            break loop;
          }
        }
        else // SKU Numbers - They are handled like this so that we can deal with multiple occurences with one alert
          illegalSkus.push(i);
      }
      else if (backgroundColours[i][j] === YELLOW) // If cell colour is Yellow
      {
        numAlerts++; // Counting the number of alerts is used to check when an export is completely with no alert messages needed

        if (i === 0) // The header row
        {
          if (j === 1) // PO Number
            sideBarMessages.push('PO number is blank, Adagio will generate the next number in sequence.')
          else if (j === 3) // Order Description
            sideBarMessages.push('Order description is too long, Adagio will chop the string at 40 characters.')
        }
        else
          if (!sideBarMessages.includes('Atleast one of your comments is too long, Adagio will chop the string at 75 characters.'))
            sideBarMessages.push('Atleast one of your comments is too long, Adagio will chop the string at 75 characters.');
      }
    }
  }

  var htmlTemplate = HtmlService.createTemplateFromFile('YellowAlert')

  if (sideBarMessages.length !== 0) // Open the sidebar
  {
    htmlTemplate.yellow_alert_1 = sideBarMessages[0]
    htmlTemplate.yellow_alert_2 = sideBarMessages[1]
    htmlTemplate.yellow_alert_3 = sideBarMessages[2]
    ui.showSidebar(htmlTemplate.evaluate().setTitle('Yellow Highlights'))
  }
  else // Close the sidebar
    ui.showSidebar(HtmlService.createHtmlOutput("<script>google.script.host.close();</script>")) // Close the sidebar

  if (response !== undefined) // There were no Red alerts (response is undefined because no red alerts means no message boxes popped up form the user)
  {
    if (response === cancel)
    {
      clearExportSheet(values, range, sheet, spreadsheet)
      ui.showSidebar(HtmlService.createHtmlOutput("<script>google.script.host.close();</script>")) // Close the sidebar
    }
    else
      sheet.getRange(1, col).activate();
  }
  else if (numAlerts !== 0) // There were red or yellow alerts
  {
    if (illegalSkus.length !== 0)
    {
      if (illegalSkus.length > 1) // Handle multiple skus here
      {
        response = ui.alert('⚠️ Change Required ⚠️', "There are multiple SKUs that need to be changed. " +
            "Please add a SKU by using the paste command (Ctrl + v) or typing the PNT SKU into column B at the appropriate rows.\n\n" +
            'Click OK, if you want to make the appropriate changes, otherwise click Cancel to clear the page.',
            ok_Cancel_ButtonSet)
      }
      else // There is only 1 illegal sku
      {
        if (values[illegalSkus[0]][1].toString() === 'SKU_NOT_FOUND')
          response = ui.alert('⚠️ Change Required ⚠️', 'Please add a SKU for ' + values[illegalSkus[0]][3].toString() + 
            ', by using the paste command (Ctrl + v) or typing the PNT SKU into column B at the appropriate row.\n\n Adagio will only accept valid SKU numbers.\n\n' + 
            'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
            ok_Cancel_ButtonSet)
        else if (values[illegalSkus[0]][1].toString() === 'NEW_ITEM_ADDED')
          response = ui.alert('⚠️ Change Required ⚠️', 'Please add a SKU for ' + values[illegalSkus[0]][3].toString() + 
            ', by using the paste command (Ctrl + v) or typing the PNT SKU into column B at the appropriate row.\n\n Adagio will only accept valid SKU numbers.\n\n' + 
            'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
            ok_Cancel_ButtonSet)
        else if (values[illegalSkus[0]][1].toString().includes(' - create new?'))
        {
          var description = values[illegalSkus[0]][3].toString().split('SKU not in Adagio. ').pop();

          response = ui.alert('⚠️ Change Required ⚠️', 'Please add a SKU for ' + description + 
            ', by using the paste command (Ctrl + v) or typing the PNT SKU into column B at the appropriate row.\n\n Adagio will only accept valid SKU numbers.\n\n' + 
            'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
            ok_Cancel_ButtonSet)
        }
        else
        {
          var description = values[illegalSkus[0]][3].toString().split('SKU not in Adagio. ').pop();
          
          response = ui.alert('⚠️ Change Required ⚠️', 'The SKU appears to be invalid. Please add a SKU for ' + description + 
            ', by using the paste command (Ctrl + v) or typing the PNT SKU into column B at the appropriate row.\n\n Adagio will only accept valid SKU numbers.\n\n' + 
            'Click OK, if you want to make the appropriate change, otherwise click Cancel to clear the page.',
            ok_Cancel_ButtonSet)
        }
      }

      if (response === cancel)
      {
        clearExportSheet(values, range, sheet, spreadsheet)
        ui.showSidebar(HtmlService.createHtmlOutput("<script>google.script.host.close();</script>")) // Close the sidebar
      }
      else
        sheet.getRange(illegalSkus[0] + 1, 2).activate();
    }
    else // Yellow Alerts only
    {
      clearExportSheet(values, range, sheet, spreadsheet)
      downloadButton(sheet, ui)
    }
  }
  else // There are no alerts
  {
    clearExportSheet(values, range, sheet, spreadsheet)
    downloadButton(sheet, ui)
  }
}

/**
 * This function launches a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of the export data being downloaded.
 * 
 * @param {Sheet} sheet : The active sheet that the export is being done from.
 * @param {Ui} ui : A user interface object.
 * @author Jarren Ralf
 */
function downloadButton(sheet, ui)
{
  var htmlTemplate = HtmlService.createTemplateFromFile('DownloadButton');
  htmlTemplate.sheetName = 'Last Export ' + sheet.getSheetName().split('Export ', 2)[1]; // This is the parameter that will be handled by the doGet() function
  var html = htmlTemplate.evaluate().setWidth(250).setHeight(50)
  ui.showModalDialog(html, 'Export');
}

/**
 * This function creates and downloads a csv document, which shows up in the user's download bar.
 * 
 * @param {String} sheetName : The name of the sheet that the export data is being generated from.
 * @return {TextOutput} Returns a text output in a csv format with a csv mime type.
 * @author Jarren Ralf
 */
function downloadCsvFile(sheetName)
{
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
  const data = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn())

  for (var row = 0, csv = ""; row < data.length; row++)
  {
    for (var col = 0; col < data[row].length; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < data.length - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('PurchaseOrder - Export.csv');
}

/**
 * This function creates the export data and pastes it on the export sheet.
 *
 * @param   {String}    vendorName  : The name of the vendor
 * @param {Object[][]}    poData    : The imported purchase order data
 * @param {Sheet}       exportSheet : The export sheet
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
 * @param {Sheet}         Sheet     : The sheet containing the newest purchase order information
 * @return {Object[][]}   output    : The po data that is created for the user
 * @author Jarren Ralf
 */
function exportInfo(vendorName, poData, exportSheet, isRefresh, spreadsheet, sheet)
{
  if (vendorName === 'Grundens')
    var output = grundens(poData, vendorName, exportSheet, isRefresh, spreadsheet)
  else if (vendorName === 'Helly Hansen')
    var output = hellyHansen(poData, vendorName, exportSheet, isRefresh, spreadsheet, sheet) // sheet is required becasue sheetName is the PO number
  else if (vendorName === 'Xtratuf')
    var output = xtratuf(poData, vendorName, exportSheet, isRefresh, spreadsheet)
  else if (vendorName === 'Yeti')
    var output = yeti(poData, vendorName, exportSheet, isRefresh, spreadsheet)
  else
    Browser.msgBox('The vendor: ' + vendorName + ' is not supported in this Spreadsheet currently.')

  if (output.length !== 0)
  {
    const horizontalAlignments = new Array(output.length).fill(['left', 'left', 'left', 'left', 'right']);

    if (isRefresh) // The user has clicked the refresh button which generates the data again (typically skipping the headers)
      var row = 3
    else
    {
      var row = 1;
      horizontalAlignments[0] = ['left', 'center', 'center', 'left', 'center'];
      horizontalAlignments[1] = ['left', 'left', 'left', 'left', 'center'];
    }
    
    exportSheet.getRange('A3').activate(); // Select the data that needs to be exported
    exportSheet.getRange(row, 1, exportSheet.getMaxRows() - row + 1, output[0].length).clearContent();
    exportSheet.getRange(row, 1, output.length, output[0].length).activate().setNumberFormat('@').setFontColor('black').setFontFamily('Arial')
      .setFontLine('none').setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setHorizontalAlignments(horizontalAlignments).setValues(output)
  }

  return output;
}

/**
 * This function adds the PNT Description to the vendors data for the sku that was just added by the user.
 * 
 * @param {Number}       row : The row number in the vendor sheet where the PNT Description is being updated
 * @param {Number}       col : The PNT SKU column
 * @param {String}       sku : The new SKU given by the user
 * @param {Sheet}      sheet : The vendor sheet
 * @param {String} sheetName : The vendor name
 * @return {String} Returns the item description from the inventory csv
 * @author Jarren Ralf
 */
function getAdagioDescription(row, col, sku, sheet, sheetName)
{
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString()); // PNT inventory data
  var skuNotInAdagio = getSkuNotFoundMessage(row, sheet, sheetName); // If the typed sku is not found, then display a message in the description column
  Logger.log(sku)
  sku = sku.toString().split(' - ', 1)[0];

  if (sku !== 'NEW_ITEM_ADDED' && sku !== 'SKU_NOT_FOUND')
  {
    var newSku = sku + ' - create new?';
    skuNotInAdagio = 'SKU not in Adagio. ' + skuNotInAdagio;
  }
  else
    var newSku = sku;

  const data = [[newSku, skuNotInAdagio]]

  for (var j = 1; j < csvData.length; j++)
  {
    if (sku == csvData[j][6]) // Match the SKUs
    {
      data[0] = [sku, csvData[j][1]] // // Add the adagio description
      break;
    }
  }

  sheet.getRange(row, col, 1, 2).setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
    .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(data);

  return data[0][1]; // Item description
}

/**
 * This function adds the PNT Descriptions to the vendors data for the skus that were just added by the user.
 * 
 * @param {Number}       row : The row number in the vendor sheet where the PNT Description is being updated
 * @param {Range[][]}  range : The range containing the pnt skus and descriptions that are being edited
 * @param {Sheet}      sheet : The vendor sheet
 * @param {String} sheetName : The vendor name
 * @author Jarren Ralf
 */
function getAdagioDescriptions(row, range, sheet, sheetName)
{
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  var sku, description;

  var items = range.getValues().map((item, i) => {
    sku = item[0].split(' - ', 1)[0];
    description = getSkuNotFoundMessage(row + i, sheet, sheetName) // If the typed sku is not found, then display a message in the description column

    if (sku !== 'NEW_ITEM_ADDED' && sku !== 'SKU_NOT_FOUND')
    {
      item[0] = sku + ' - create new?';
      item[1] = 'SKU not in Adagio. ' + description;
    }

    for (var j = 1; j < csvData.length; j++)
    {
      if (sku == csvData[j][6]) // Match the SKUs
      {
        item = [sku, csvData[j][1]] // // Add the adagio description
        break;
      }
    }

    return item;
  })

  range.setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal')
    .setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(items);
}

/**
 * This function produces a proposed SKU for a particular Xtratuf product.
 * 
 * @param {String} stockNumber : The stock number of a particular Xtratuf item.
 * @param {String}     size    : The size of a particular Xtratuf item.
 * @return {String} The proposed Xtratuf SKU.
 * @author Jarren Ralf
 */
function getProposedNewXtratufSKU(stockNumber, size)
{
  var sku = '';

  switch (stockNumber.toString().length)
  {
    case 1:
      sku += '1011011' + stockNumber.toString() + size.toString().replace(/\./g,'')
      break;
    case 2:
      sku += '101101' + stockNumber.toString() + size.toString().replace(/\./g,'')
      break;
    case 3:
      sku += '10110' + stockNumber.toString() + size.toString().replace(/\./g,'')
      break;
    case 4:
      sku += '1011' + stockNumber.toString() + size.toString().replace(/\./g,'')
      break;
    case 5:
      sku += '101' + stockNumber.toString() + size.toString().replace(/\./g,'')
      break;
    case 6:
      sku += '10' + stockNumber.toString() + size.toString().replace(/\./g,'')
      break;
    default:
      sku += '1' + stockNumber.toString() + size.toString().replace(/\./g,'')
  }

  return sku;
}

/**
 * This function produces a proposed SKU for a particular Yeti product.
 * 
 * @param {String}    category : The category of a particular Yeti item.
 * @param {String} description : The description of a particular Yeti item.
 * @return {String} The proposed Yeti SKU.
 * @author Jarren Ralf
 */
function getProposedNewYetiSKU(category, description)
{
  description = description.toUpperCase().replace(' / ', '/').replace('/', ' / ').replace(' w / ', ' w/ ').split(' '); // Reformat the Yeti description and split it

  if (description[0] === 'INTL') // Remove the 'INTL' from the beginning of the array
    description.shift();

  const clr = getYetiColourAbbreviation(description); // This is the colour abbreviation for our Yeti skus
  var sku = '';

  switch (category.toUpperCase())
  {
    case 'BOOMER DOG BOWL':
      sku = '8015201' + description[1] + clr;
      break;
    case 'BOTTLE ACCESSORY':
      sku = '8015RBSL' + description[3] + clr;
      break;
    case 'BOTTLES':
      switch (description[1]) // Size
      {
        case 'JR':
          sku = '8015RBJ' + description[2].substring(description[2].length - 2, 0) + clr; // Remove the units 'oz' (the last 2 characters)
          break;
        case 'HALF':
          sku = '8015RJHG' + clr;
          break;
        case 'ONE':
          sku = '8015RJG' + clr;
          break;
        default:
          switch (description[3])
          {
            case 'HOTSHOT':
              sku = '8015RHB' + description[1].substring(description[1].length - 2, 0) + clr; // Remove the units 'oz' (the last 2 characters)
              break;
            default:
              sku = '8015RB' + description[1].substring(description[1].length - 2, 0) + clr; // Remove the units 'oz' (the last 2 characters)
          }
      }
      break;
    case 'CARGO':
      switch (description[0]) // Style Name
      {
        case 'CAMINO':
          sku = '2186CAM35' + clr;
          break;
        case 'PANGA':
          switch (description[2]) // Type
          {
            case 'DUFFEL':
              sku = '2PANGA' + description[3] + clr;
              break;
            case 'SUBMERSIBLE': // Backpack
              sku = '2PANGA' + description[1] + clr;
              break;
          }
          break;
      }
      break;
    case 'DAYTRIP':
      sku = '8015DT' + description[2] + clr;
      break;
    case 'DRINKWARE':
      switch (description[4]) // Size
      {
        case 'MUG':
          sku = '8015RM' + description[1].substring(description[1].length - 2, 0) + clr; // Remove the units 'oz' (the last 2 characters)
          break;
        case 'LOWBALL':
          sku = '8015RL' + description[1].substring(description[1].length - 2, 0) + clr;
          break;
        case 'TUMBLER':
          sku = '8015RT' + description[1].substring(description[1].length - 2, 0) + clr;
          break;
        case 'STRAW':
          sku = '8015RSC' + description[1].substring(description[1].length - 2, 0) + clr;
          break;
      }
      break;
    case 'DRINKWARE ACCESSORIES':
      switch (description[description.length - 1]) // Type
      {
        case 'HANDLE':
          switch (description[4]) // Type
          {
            case 'LOWBALL':
              sku = '8015RLHL' + description[1].substring(description[1].length - 2, 0) // Remove the units 'oz' (the last 2 characters)
              break;
            case 'TUMBLER':
              sku = '8015RTHL' + description[1].substring(description[1].length - 2, 0)
              break;
          }
          break;
        case 'LID':
          switch (description[description.length - 2]) // Type
          {
            case 'MAGSLIDER':
              switch (description[5]) // Type
              {
                case 'LID':
                  sku = '8015RTLID' + description[1].substring(description[1].length - 2, 0) // Remove the units 'oz' (the last 2 characters)
                  break;
                case 'TUMBLER':
                  sku = '8015WINELID' + description[1].substring(description[1].length - 2, 0) // Remove the units 'oz' (the last 2 characters)
                  break;
                default:
                  sku = '8015RTLID' + description[1].substring(description[1].length - 2, 0) + description[5].substring(description[5].length - 2, 0)
              }
              break;
            case 'STRAW':
              sku = '8015RTSTRAWLID' + description[2].substring(description[2].length - 2, 0) 
              break;
          }
          break;
        case 'CAP':
          sku = '8015RT' + description[1].substring(description[1].length - 2, 0) + clr;
          break;
        case 'MOUNT':
          switch (description[1]) // Size
          {
            case 'HALF':
              sku = '8015RJHGMOUNT';
              break;
            case 'ONE':
              sku = '8015RJGMOUNT';
              break;
          }
          break;
      }
      break;
    case 'HOPPER SERIES':
      sku = '2186' + (description[1].length === 3) ? '0' + description[1] + clr : description[1] + clr;
      break;
    case 'LOADOUT BUCKET':
      switch (description[2]) // Style Type
      {
        case 'UTILITY':
          sku = '2BUCKETGEAR' + clr;
          break;
        case 'LID':
          sku = '2BUCKETLID' + clr;
          break;
        default:
          sku = '2BUCKET' + clr;
      }
      break;
    case 'LOADOUT GOBOX':
      sku = '2GOBOX' + description[2] + clr;
      break;
    case 'OUTDOOR LIVING':
      switch (description[0]) // Style Type
      {
        case 'LOWLANDS':
          sku = '80150195' + clr;
          break;
        case 'TRAILHEAD':
          sku = '80150196'  + clr;
          break;
      }
      break;
    case 'ROADIE SERIES':
      sku = '2186ROADIE' + description[1] + clr;
      break;
    case 'SILO':
      sku = '8015SILO' + description[1] + description[2] + clr;
      break;
    case 'TANK SERIES':
      switch (description[2]) // Style Type
      {
        case 'LID':
          sku = '2186TANK' + description[1] + 'LID' + clr;
          break;
        default:
          sku = '2186TANK' + description[1] + clr
      }
      break;
    case 'TUNDRA SERIES':
      sku = '2186' + ((description[1].length == 4) ? description[1] + clr : (description[1].length == 3) ? '0' + description[1] + clr : '00' + description[1] + clr);
      break;
    case 'YETI ICE':
      sku = 'YETIICE' + description[2] + clr;
      break;
  }

  return sku;
}

/**
 * This fucntion gets the message that will populate the PNT Description in the given vendor database when the particular sku is not found in Adagio.
 * 
 * @param {Number}       row : The row number that the user is editing on the current sheet. 
 * @param {Sheet}      sheet : The sheet that the current user is editing.
 * @param {String} sheetName : The name of the sheet that the current user is editing.
 * @return {String} The message that will replace the PNT description for the case when a particular sku is not found in Adagio.
 * @author Jarren Ralf
 */
function getSkuNotFoundMessage(row, sheet, sheetName)
{
  const lastCol = sheet.getLastColumn()
  const  header = sheet.getSheetValues(  1, 1, 1, lastCol)[0]
  const    item = sheet.getSheetValues(row, 1, 1, lastCol)[0]

  if (sheetName === 'Grundens')
  {
    const grunDescrip = header.indexOf('Name');
    const grunColour = header.indexOf('Color');
    const grunSize = header.indexOf('Size 1');
    const grunSubCat = header.indexOf('Subcategory');
    const grunCat = header.indexOf('Category');
    const grunPrice = header.indexOf('Price 1 USD');
    const grunStyleNum = header.indexOf('Style Number');
    const grunColorCode = header.indexOf('Color Code');

    const proposedNewSKU = item[grunStyleNum].toString() + ((item[grunColorCode].toString().length == 1) ? '00' + item[grunColorCode].toString() : 
                                                            (item[grunColorCode].toString().length == 2) ?  '0' + item[grunColorCode].toString() : 
                                                             item[grunColorCode].toString()) + ((item[grunSize].toString() === 'XXL') ? '2XL' : item[grunSize].toString());
    return item[grunDescrip] + ' - ' + item[grunColour]   + ' - ' + item[grunSize] + ' - ' + 
            ((item[grunSubCat] !== '') ? item[grunSubCat] + ' - ' : '- ') +
            ((item[grunCat]    !== '') ? item[grunCat]    + ' - ' : '- ') + 'Cost: $' + Number(item[grunPrice]).toFixed(2) + ' - ' + proposedNewSKU;
  }
  else if (sheetName === 'Helly Hansen')
  {
    const hellyDescription = header.indexOf('Style Name')
    const hellyColour = header.indexOf('Color Name')
    const hellySize = header.indexOf('Size')
    const hellyPrice = header.indexOf('Wholesale Price')
    const hellySKU = header.indexOf('SKU')
    const proposedNewSKU = item[hellySKU].replace(/_|-/g,'')

    return item[hellyDescription] + ' - ' + item[hellyColour] 
              + ' - ' + item[hellySize] + ' - Cost: $' + Number(item[hellyPrice]).toFixed(2) + ' - ' + proposedNewSKU;
  }
  else if (sheetName === 'Xtratuf')
  {
    const xtratufDescription = header.indexOf('Name/Nom/Description')
    const xtratufCategory = header.indexOf('Category / Catégorie')
    const xtratufColour = header.indexOf('Color / Couleur')
    const xtratufSize = header.indexOf('Sizes / Tailles')
    const xtratufPrice = header.indexOf('Purchase Price/\nPrix d\’achat ')
    const xtratufSku = header.indexOf('Stock# / Nº de nomenclature')
    const proposedNewSKU = getProposedNewXtratufSKU(item[xtratufSku], item[xtratufSize])

    return item[xtratufDescription] + ' - ' + item[xtratufCategory] + ' - ' + item[xtratufColour]
             + ' - ' + item[xtratufSize] + ' - Cost: $' + Number(item[xtratufPrice]).toFixed(2) + ' - ' + proposedNewSKU;
  }
  else if (sheetName === 'Yeti')
  {
    const yetiDescription = header.indexOf('DESCRIPTION')
    const yetiCategory = header.indexOf('CATEGORY')
    const yetiPrice = header.indexOf('DEALER PRICE')
    const proposedNewSKU = getProposedNewYetiSKU(item[yetiCategory], item[yetiDescription])

    return item[yetiDescription] + ' - ' + item[yetiCategory] + ' - Cost: $' + Number(item[yetiPrice]).toFixed(2) + ' - ' + proposedNewSKU;
  }
  else
    return "SKU not in Adagio. The " + sheetName + " sheet does not have a descriptive message programmed. Contact the spreadsheet owner to add this feature."
}

/**
 * This function figures out what the vendor is based on the imported PO data. Each vendor will have it's own set of checks.
 * 
 * @param {Object[][]} values : The imported purchase order data
 * @return {String} Returns the Vendor name.
 * @author Jarren Ralf
 */
function getVendorName(values)
{
  Logger.log(values[1][0])
  
  if (values[1][0] === 'Grundens')
    return 'Grundens';
  else if (values[0].includes("Seasons") && values[0].includes("Segmentation") && values[0].includes("Technology"))
    return 'Helly Hansen';
  else if (values[18][1] === 'XTRATUF')
    return 'Xtratuf';
  else if (values[3][4] === 'http://www.yeti.ca')
    return 'Yeti';
  else
    return 'VENDOR_NAME_NOT_DETECTED'
}

/**
 * This function retrieves the vendor number from the Vendors page.
 * 
 * @param   {String}    vendorName  : The name of the vendor
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @return {String} Returns the Vendor number.
 * @author Jarren Ralf
 */
function getVendorNumber(vendorName, spreadsheet)
{
  const vendorSheet = spreadsheet.getSheetByName('Vendors');
  const vendorData = vendorSheet.getSheetValues(2, 1, vendorSheet.getLastRow() - 1, 2);

  for (var k = 0; k < vendorData.length; k++)
    if (vendorData[k][0] === vendorName) return vendorData[k][1];
}

/**
 * This function produces an abbreviation for the colour of a Yeti product. IF the colour is undeterminable, then 
 * the colour abbreviation that is returned is blank ''.
 * 
 * @param {String[]} description : The description of a particular Yeti item (UPPERCASE) split into an array by a space ' '.
 * @return {String} The abbreviation for the given colour.
 * @author Jarren Ralf
 */
function getYetiColourAbbreviation(description)
{
  var clr = '';

  switch (description[description.length - 1]) // Take the last word in the description and assume this is the colour
  {
    case 'BLACK':
      switch (description[description.length - 2]) // Second last word
      {
        case 'CUSHION':
          clr = '';
          break;
        default:
          clr = 'BK'
      }
      break;
    case 'BLUE':
      switch (description[description.length - 2]) // Second last word
      {
        case 'NORDIC':
          clr = 'SC1'; // Seasonal Colour # 1
          break;
        case 'OFFSHORE':
          clr = 'SD'; // Seasonal Discontinued
          break;
        case 'REEF':
          clr = 'RB';
          break;
        case 'SMOKE':
          clr = 'SB';
          break;
        default:
          clr = 'BLU'
      }
      break;
    case 'CADDY':
      clr = 'CADDY';
      break;
    case 'CHARCOAL':
      clr = 'CHAR';
      break;
    case 'CHR':
      clr = 'CHR';
      break;
    case 'GRAY':
      switch (description[description.length - 2]) // Second last word
      {
        case 'STORM':
          clr = 'SG';
          break;
        default:
          clr = 'GRAY'
      }
      break;
    case 'NAVY':
    case 'NVY':
      clr = 'NAVY';
      break;
    case 'OLIVE':
      switch (description[description.length - 2]) // Second last word
      {
        case 'HIGHLANDS':
          clr = 'SD'; // Seasonal Discontinued
          break;
        default:
          clr = 'OLV'
      }
      break;
    case 'PINK':
      switch (description[description.length - 2]) // Second last word
      {
        case 'BIMINI':
        case 'SANDSTONE':
          clr = 'SD'; // Seasonal Discontinued
          break;
        case 'HARBOUR':
          clr = 'HP';
          break;
        default:
          clr = 'PINK'
      }
      break;
    case 'PURPLE':
      switch (description[description.length - 2]) // Second last word
      {
        case 'NORDIC':
          clr = 'SC2'; // Seasonal Colour # 2
          break;
        default:
          clr = 'PUR'
      }
      break;
    case 'RED':
      switch (description[description.length - 2]) // Second last word
      {
        case 'CANYON':
          clr = 'CR';
          break;
        case 'FIRESIDE':
          clr = 'FR';
          break;
        case 'HARVEST':
          clr = 'SD'; // Seasonal Discontinued
          break;
        default:
          clr = 'RED'
      }
      break;
    case 'SEAFOAM':
      clr = 'SF';
      break;
    case 'STAINLESS':
      clr = 'SS';
      break;
    case 'TAN':
      clr = 'TAN';
      break;
    case 'TAUPE':
      switch (description[description.length - 2]) // Second last word
      {
        case 'SHARPTAIL':
          clr = 'SD'; // Seasonal Discontinued
          break;
        default:
          clr = 'TAUPE'
      }
      break;
    case 'WHITE':
    case 'WHT':
      clr = 'WHT';
      break;
    case 'YELLOW':
      switch (description[description.length - 2]) // Second last word
      {
        case 'ALPINE':
          clr = 'SD'; // Seasonal Discontinued
          break;
        default:
          clr = 'YLW'
      }
      break;
    default:
      clr = '';
  }

  return clr;
}

//This is the original Grundens code that only supports 
// /**
//  * This function creates the export data for a Grundens purchase order.
//  * 
//  * @param {Object[][]}       poData : The purchase order data that was just uploaded.
//  * @param {String}         Grundens : The name of the vendor, in this casem Grundens.
//  * @param {Sheet}       exportSheet : The sheet that the data will be exported to.
//  * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
//  * @param {Spreadsheet} spreadsheet : The active spreadsheet.
//  * @return {Object[][]}      output : The data created for the export sheet.
//  * @author Jarren Ralf
//  */
// function grundens(poData, Grundens, exportSheet, isRefresh, spreadsheet)
// {
//   const grundensSheet = spreadsheet.getSheetByName('Grundens');
//   const grunData = grundensSheet.getDataRange().getValues();
//   const vendorNumber = getVendorNumber(Grundens, spreadsheet);
//   var output = [], addToGrundensData = [], totalCost = 0, proposedNewSKU;

//   // Grundens Data
//   const cost           = grunData[0].indexOf("Price 1 USD");
//   const sizing         = grunData[0].indexOf("Size 1");
//   const pntSKU         = grunData[0].indexOf("PNT SKU");
//   const colourCode     = grunData[0].indexOf("Color Code");
//   const styleNumber    = grunData[0].indexOf("Style Number");
//   const pntDescription = grunData[0].indexOf("PNT Description");

//   // Purchase Order Data
//   const qty        = poData[0].indexOf("Quantity"); 
//   const size       = poData[0].indexOf("Size");
//   const poNum      = poData[0].indexOf("Customer PO Number");
//   const style      = poData[0].indexOf("Style Number");
//   const price      = poData[0].indexOf("Price Per");
//   const color      = poData[0].indexOf("Color");
//   const colorCode  = poData[0].indexOf("Color Code");
//   const styleName  = poData[0].indexOf("Style Name");
//   const orderTags  = poData[0].indexOf("Order Tags");

//   for (var i = 1; i < poData.length; i++)
//   {
//     for (var j = 1; j < grunData.length; j++)
//     {
//       // Find the item in the Grundens data base
//       if (poData[i][style] == grunData[j][styleNumber] && Number(poData[i][colorCode]) == Number(grunData[j][colourCode]) && poData[i][size] == grunData[j][sizing])
//       {
//         totalCost += poData[i][qty]*grunData[j][cost];
//         output.push(['R', grunData[j][pntSKU], poData[i][qty], grunData[j][pntDescription], Number(grunData[j][cost]).toFixed(2)]) // Receiving Line
//         break;
//       }
//     }

//     if (j === grunData.length) // The item(s) that were ordered were not found on the Grundens data sheet
//     {
//       proposedNewSKU = poData[i][style].toString() + ((poData[i][colorCode].toString().length == 1) ? '00' + poData[i][colorCode].toString() : 
//                                                       (poData[i][colorCode].toString().length == 2) ?  '0' + poData[i][colorCode].toString() : 
//                                                        poData[i][colorCode].toString()) + poData[i][size]
//       addToGrundensData.push([poData[i][style], poData[i][orderTags], toProper(poData[i][color]), null, null, toProper(poData[i][styleName]), poData[i][colorCode], 
//         null, null, null, null, null, null, null, null, null, poData[i][size], null, Number(poData[i][price]).toFixed(2), null, 'NEW_ITEM_ADDED', 
//         toProper(poData[i][styleName]) + ' - ' + toProper(poData[i][color]) + ' - '  + poData[i][size] + ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
//       output.push(['R', 'NEW_ITEM_ADDED', poData[i][qty], toProper(poData[i][styleName]) + ' - ' + toProper(poData[i][color]) + ' - '  + poData[i][size] + 
//         ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)])
//     }
//   }

//   if (addToGrundensData.length !== 0)
//   {
//     grundensSheet.showSheet().getRange(grundensSheet.getLastRow() + 1, 1, addToGrundensData.length, addToGrundensData[0].length).activate().setHorizontalAlignment('left')
//       .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
//       .setValues(addToGrundensData)
//     spreadsheet.toast('Add PNT SKUs to the Grundens data.', '⚠️ New Grundens Items ⚠️', 30)
//   }

//   if (output.length !== 0 && !isRefresh)
//   {
//     var currentEditor = exportSheet.getSheetValues(2, 5, 1, 1)[0][0];
//     var po = reformatPoNumber(poData[1][poNum].toString());

//     if (currentEditor === '')
//       currentEditor = 'Someone is currently editing this PO'

//     output.unshift(
//       ['H', po, vendorNumber, 'Type your order description in this cell. (40 characters max)', Grundens], // Header line
//       ['C', 'Type your order comments in this cell. (75 characters max)', null, null, currentEditor] // Comment Line
//     )
//   }
//   return output;
// }

/**
 * This function creates the export data for a Grundens purchase order. Is able to import multiple purchase orders.
 * 
 * @param {Object[][]}       poData : The purchase order data that was just uploaded.
 * @param {String}         Grundens : The name of the vendor, in this casem Grundens.
 * @param {Sheet}       exportSheet : The sheet that the data will be exported to.
 * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @return {Object[][]}      output : The data created for the export sheet.
 * @author Jarren Ralf
 */
function grundens(poData, Grundens, exportSheet, isRefresh, spreadsheet)
{
  const grundensSheet = spreadsheet.getSheetByName('Grundens');
  const grunData = grundensSheet.getDataRange().getValues();
  var addToGrundensData = [], proposedNewSKU;

  // Grundens Data
  const cost           = grunData[0].indexOf("Price 1 USD");
  const sizing         = grunData[0].indexOf("Size 1");
  const pntSKU         = grunData[0].indexOf("PNT SKU");
  const colourCode     = grunData[0].indexOf("Color Code");
  const styleNumber    = grunData[0].indexOf("Style Number");
  const pntDescription = grunData[0].indexOf("PNT Description");

  // Purchase Order Data
  const qty        = poData[0].indexOf("Quantity"); 
  const size       = poData[0].indexOf("Size");
  const poNum      = poData[0].indexOf("Customer PO Number");
  const style      = poData[0].indexOf("Style Number");
  const price      = poData[0].indexOf("Price Per");
  const color      = poData[0].indexOf("Color");
  const colorCode  = poData[0].indexOf("Color Code");
  const styleName  = poData[0].indexOf("Style Name");
  const orderTags  = poData[0].indexOf("Order Tags");
  const po = reformatPoNumber(poData[1][poNum].toString());
  const vendorNumber = getVendorNumber(Grundens, spreadsheet);
  var currentEditor = exportSheet.getSheetValues(2, 5, 1, 1)[0][0];
  
  if (currentEditor === '')
    currentEditor = 'Someone is currently editing this PO';

  const output = (isRefresh) ? [] : [['H', po, vendorNumber, 'Type your order description in this cell. (40 characters max)', Grundens],
                                     ['C', 'Type your order comments in this cell. (75 characters max)', null, null, currentEditor]]

  for (var i = 1; i < poData.length; i++)
  {
    for (var j = 1; j < grunData.length; j++)
    {
      if (poData[i][style] == grunData[j][styleNumber] && Number(poData[i][colorCode]) == Number(grunData[j][colourCode]) && poData[i][size] == grunData[j][sizing])
      {
        if (isNewPurchaseOrder(poData, i, poNum)) // If a new Purchase Order is detected, then we need to add a new header
          output.push(['H', po, vendorNumber, 'Type your order description in this cell. (40 characters max)', Grundens])

        output.push(['R', grunData[j][pntSKU], poData[i][qty], grunData[j][pntDescription], Number(grunData[j][cost]).toFixed(2)]) // Receiving Line
        break;
      }
    }

    if (j === grunData.length) // The item(s) that were ordered were not found on the Grundens data sheet
    {
      proposedNewSKU = poData[i][style].toString() + ((poData[i][colorCode].toString().length == 1) ? '00' + poData[i][colorCode].toString() : 
                                                      (poData[i][colorCode].toString().length == 2) ?  '0' + poData[i][colorCode].toString() : 
                                                       poData[i][colorCode].toString()) + poData[i][size]
      addToGrundensData.push([poData[i][style], poData[i][orderTags], toProper(poData[i][color]), null, null, toProper(poData[i][styleName]), poData[i][colorCode], 
        null, null, null, null, null, null, null, null, null, poData[i][size], null, Number(poData[i][price]).toFixed(2), null, 'NEW_ITEM_ADDED', 
        toProper(poData[i][styleName]) + ' - ' + toProper(poData[i][color]) + ' - '  + poData[i][size] + ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
      output.push(['R', 'NEW_ITEM_ADDED', poData[i][qty], toProper(poData[i][styleName]) + ' - ' + toProper(poData[i][color]) + ' - '  + poData[i][size] + 
        ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)])
    }
  }

  if (addToGrundensData.length !== 0) // Add items to the Grundens data that were on the PO but not in the Grundens data
  {
    grundensSheet.showSheet().getRange(grundensSheet.getLastRow() + 1, 1, addToGrundensData.length, addToGrundensData[0].length).activate().setHorizontalAlignment('left')
      .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
      .setValues(addToGrundensData)
    spreadsheet.toast('Add PNT SKUs to the Grundens data.', '⚠️ New Grundens Items ⚠️', 30)
  }

  return output;
}

// /**
//  * This function creates the export data for a Helly Hansen purchase order.
//  * 
//  * @param {Object[][]}       poData : The purchase order data that was just uploaded.
//  * @param {String}      HellyHansen : The name of the vendor, in this case Helly Hansen.
//  * @param {Sheet}       exportSheet : The sheet that the data will be exported to.
//  * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
//  * @param {Spreadsheet} spreadsheet : The active spreadsheet.
//  * @param {Sheet}             sheet : The sheet containing the newest purchase order information
//  * @return {Object[][]}      output : The data created for the export sheet.
//  * @author Jarren Ralf
//  */
// function hellyHansen(poData, HellyHansen, exportSheet, isRefresh, spreadsheet, sheet)
// {
//   const hellySheet = spreadsheet.getSheetByName('Helly Hansen');
//   const hellyData = hellySheet.getDataRange().getValues();
//   const vendorNumber = getVendorNumber(HellyHansen, spreadsheet);
//   var output = [], addToHellyHansenData = [], proposedNewSKU;

//   // Helly Hansen Data
//   const cost           = hellyData[0].indexOf("Wholesale Price");
//   const SKU            = hellyData[0].indexOf("SKU");
//   const pntSKU         = hellyData[0].indexOf("PNT SKU");
//   const pntDescription = hellyData[0].indexOf("PNT Description");

//   // Purchase Order Data
//   const qty          = poData[0].indexOf("Quantity Requested");
//   const sku          = poData[0].indexOf('SKU')
//   const styleName    = poData[0].indexOf("Style Name");
//   const style        = poData[0].indexOf("Style Number");
//   const color        = poData[0].indexOf("Color Name");
//   const colorCode    = poData[0].indexOf("Color Code");
//   const size         = poData[0].indexOf("Size");
//   const altSize      = poData[0].indexOf("Alt Size");
//   const upc          = poData[0].indexOf("UPC");
//   const wholsaleCost = poData[0].indexOf("Wholesale Price");
//   const price        = poData[0].indexOf("Retail Price");
//   const retailDate   = poData[0].indexOf("Retail Date");
//   const availableQty = poData[0].indexOf("QuantityAvailable");
//   const category     = poData[0].indexOf("Category");
//   const certify      = poData[0].indexOf("Certification");
//   const concept      = poData[0].indexOf("Concept");
//   const fitting      = poData[0].indexOf("Fitting");
//   const gender       = poData[0].indexOf("Gender");
//   const group        = poData[0].indexOf("Product Group");
//   const seasons      = poData[0].indexOf("Seasons");
//   const segmentation = poData[0].indexOf("Segmentation");
//   const technology   = poData[0].indexOf("Technology");
//   const status       = poData[0].indexOf("Status");
//   const colour       = poData[0].indexOf("Color");

//   for (var i = 1; i < poData.length; i++)
//   {
//     for (var j = 1; j < hellyData.length; j++)
//     {
//       // Find the item in the Helly Hansen data base
//       if (poData[i][sku] == hellyData[j][SKU])
//       {
//         output.push(['R', hellyData[j][pntSKU], poData[i][qty], hellyData[j][pntDescription], Number(hellyData[j][cost]).toFixed(2)]) // Receiving Line
//         break;
//       }
//     }

//     if (j === hellyData.length) // The item(s) that were ordered were not found on the Helly Hansen data sheet
//     {
//       proposedNewSKU = poData[i][sku].replace(/_|-/g,'')
//       addToHellyHansenData.push([poData[i][styleName], poData[i][style], poData[i][color], poData[i][colorCode], null, poData[i][size], poData[i][altSize], 
//         poData[i][upc], poData[i][sku], Number(poData[i][wholsaleCost]).toFixed(2), Number(poData[i][price]).toFixed(2), poData[i][retailDate], null, poData[i][availableQty], 
//         poData[i][category], poData[i][certify], poData[i][concept], poData[i][fitting], poData[i][gender], poData[i][group], poData[i][seasons], 
//         poData[i][segmentation], poData[i][technology], poData[i][status], poData[i][colour], 'NEW_ITEM_ADDED', poData[i][styleName] + ' - ' + poData[i][color] 
//         + ' - '  + poData[i][size] + ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
//       output.push(['R', 'NEW_ITEM_ADDED', poData[i][qty], poData[i][styleName] + ' - ' + poData[i][color] + ' - '  + poData[i][size] + 
//         ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)]) // Receiving Line
//     }
//   }

//   if (addToHellyHansenData.length !== 0) // Add items to the Helly Hansen data that were on the PO but not in the Helly Hansen data
//   {
//     hellySheet.showSheet().getRange(hellySheet.getLastRow() + 1, 1, addToHellyHansenData.length, addToHellyHansenData[0].length).activate().setHorizontalAlignment('left')
//       .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
//       .setValues(addToHellyHansenData)
//     spreadsheet.toast('Add PNT SKUs to the Helly Hansen data.', '⚠️ New Helly Hansen Items ⚠️', 30)
//   }

//   if (output.length !== 0 && !isRefresh)
//   {
//     var currentEditor = exportSheet.getSheetValues(2, 5, 1, 1)[0][0];
//     var po = (sheet === undefined) ? '' : reformatPoNumber(sheet.getSheetName()); // (sheet === undefined) occurs when the refresh button is clicked when the export page is blank

//     if (currentEditor === '')
//       currentEditor = 'Someone is currently editing this PO'

//     output.unshift(
//       ['H', po, vendorNumber, 'Type your order description in this cell. (40 characters max)', HellyHansen], // Header line
//       ['C', 'Type your order comments in this cell. (75 characters max)', null, null, currentEditor] // Comment Line
//     )
//   }
//   return output;
// }

/**
 * This function creates the export data for a Helly Hansen purchase order.
 * 
 * @param {Object[][]}       poData : The purchase order data that was just uploaded.
 * @param {String}      HellyHansen : The name of the vendor, in this case Helly Hansen.
 * @param {Sheet}       exportSheet : The sheet that the data will be exported to.
 * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {Sheet}             sheet : The sheet containing the newest purchase order information
 * @return {Object[][]}      output : The data created for the export sheet.
 * @author Jarren Ralf
 */
function hellyHansen(poData, HellyHansen, exportSheet, isRefresh, spreadsheet, sheet)
{
  const hellySheet = spreadsheet.getSheetByName('Helly Hansen');
  const hellyData = hellySheet.getDataRange().getValues();
  var addToHellyHansenData = [], proposedNewSKU;

  // Helly Hansen Data
  const cost           = hellyData[0].indexOf("Wholesale Price");
  const SKU            = hellyData[0].indexOf("SKU");
  const pntSKU         = hellyData[0].indexOf("PNT SKU");
  const pntDescription = hellyData[0].indexOf("PNT Description");

  // Purchase Order Data
  const qty          = poData[0].indexOf("Quantity Requested");
  const sku          = poData[0].indexOf('SKU')
  const styleName    = poData[0].indexOf("Style Name");
  const style        = poData[0].indexOf("Style Number");
  const color        = poData[0].indexOf("Color Name");
  const colorCode    = poData[0].indexOf("Color Code");
  const size         = poData[0].indexOf("Size");
  const altSize      = poData[0].indexOf("Alt Size");
  const upc          = poData[0].indexOf("UPC");
  const wholsaleCost = poData[0].indexOf("Wholesale Price");
  const price        = poData[0].indexOf("Retail Price");
  const retailDate   = poData[0].indexOf("Retail Date");
  const availableQty = poData[0].indexOf("QuantityAvailable");
  const category     = poData[0].indexOf("Category");
  const certify      = poData[0].indexOf("Certification");
  const concept      = poData[0].indexOf("Concept");
  const fitting      = poData[0].indexOf("Fitting");
  const gender       = poData[0].indexOf("Gender");
  const group        = poData[0].indexOf("Product Group");
  const seasons      = poData[0].indexOf("Seasons");
  const segmentation = poData[0].indexOf("Segmentation");
  const technology   = poData[0].indexOf("Technology");
  const status       = poData[0].indexOf("Status");
  const colour       = poData[0].indexOf("Color");
  const po           = poData[0].indexOf("PO #");
  const vendorNumber = getVendorNumber(HellyHansen, spreadsheet);
  var currentEditor = exportSheet.getSheetValues(2, 5, 1, 1)[0][0];

  if (currentEditor === '')
      currentEditor = 'Someone is currently editing this PO'

  const output = (isRefresh) ? [] : [['H', reformatPoNumber(poData[1][po]), vendorNumber, 'Type your order description in this cell. (40 characters max)', HellyHansen],
                                     ['C', 'Type your order comments in this cell. (75 characters max)', null, null, currentEditor]]


  for (var i = 1; i < poData.length; i++)
  {
    for (var j = 1; j < hellyData.length; j++)
    {
      // Find the item in the Helly Hansen data base
      if (poData[i][sku] == hellyData[j][SKU])
      {
        output.push(['R', hellyData[j][pntSKU], poData[i][qty], hellyData[j][pntDescription], Number(hellyData[j][cost]).toFixed(2)]) // Receiving Line
        break;
      }
    }

    if (j === hellyData.length) // The item(s) that were ordered were not found on the Helly Hansen data sheet
    {
      proposedNewSKU = poData[i][sku].replace(/_|-/g,'')
      addToHellyHansenData.push([poData[i][styleName], poData[i][style], poData[i][color], poData[i][colorCode], null, poData[i][size], poData[i][altSize], 
        poData[i][upc], poData[i][sku], Number(poData[i][wholsaleCost]).toFixed(2), Number(poData[i][price]).toFixed(2), poData[i][retailDate], null, poData[i][availableQty], 
        poData[i][category], poData[i][certify], poData[i][concept], poData[i][fitting], poData[i][gender], poData[i][group], poData[i][seasons], 
        poData[i][segmentation], poData[i][technology], poData[i][status], poData[i][colour], 'NEW_ITEM_ADDED', poData[i][styleName] + ' - ' + poData[i][color] 
        + ' - '  + poData[i][size] + ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
      output.push(['R', 'NEW_ITEM_ADDED', poData[i][qty], poData[i][styleName] + ' - ' + poData[i][color] + ' - '  + poData[i][size] + 
        ' - - - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)]) // Receiving Line
    }
  }

  if (addToHellyHansenData.length !== 0) // Add items to the Helly Hansen data that were on the PO but not in the Helly Hansen data
  {
    hellySheet.showSheet().getRange(hellySheet.getLastRow() + 1, 1, addToHellyHansenData.length, addToHellyHansenData[0].length).activate().setHorizontalAlignment('left')
      .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
      .setValues(addToHellyHansenData)
    spreadsheet.toast('Add PNT SKUs to the Helly Hansen data.', '⚠️ New Helly Hansen Items ⚠️', 30)
  }

  return output;
}

/**
 * This function updates the inventory from a csv file. It also hides all of the non-primary sheets (since this function is run daily on a trigger).
 * 
 * @author Jarren Ralf
 */
function importInventory()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const sheets = spreadsheet.getSheets();
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString()); // PNT inventory data
  var header = csvData.shift(); // Remove the header (we don't want the header to be sorted into the middle of the array)
  const activeItems = csvData.filter(item => item[10] === 'A').sort(sortByCategories) // Remove the inactive items and sort by the categories
  var sheetName; // Will be used when hiding sheets

  const recentlyCreatedItems = activeItems.map(val => {
    d = val[7].split('.');                           // Split the date at the "."
    val[7] = new Date(d[2],d[1] - 1,d[0]).getTime(); // Convert the date string to a Number for convenient sorting purposes
    return val;
  }).sort(sortByCreatedDate).map(c => [c[0], c[1], null, c[2], c[3], c[4]]); // Sort by the creative date, then only keep the desired columns

  const numRows = activeItems.unshift(header); // Add the header back to the top of the array
  header = [header[0], header[1], null, header[2], header[3], header[4]]
  recentlyCreatedItems.unshift(header) // Add a header to the recently created items
  const data = activeItems.map(c => [c[0], c[1], null, c[2], c[3], c[4]]); // Remove the On Transfer Sheet, Comments 3 (Categories), and Active Item columns
  spreadsheet.getSheetByName('Inventory').clearContents().getRange(1, 1, numRows, data[0].length).setNumberFormat('@').setValues(data);
  spreadsheet.getSheetByName('Recent').clearContents().getRange(1, 1, numRows, data[0].length).setNumberFormat('@').setValues(recentlyCreatedItems);

  // Since import inventory will run on a trigger daily, the opportunity is taken to hide the sheets we don't want displayed
  sheets.forEach(sheet => {
    sheetName = sheet.getSheetName();

    if (sheetName !== 'Search 1' && sheetName !== 'Search 2' && sheetName !== 'Export 1' && sheetName !== 'Export 2') // Keep these sheets displayed, hide the rest
      sheet.hideSheet();
  })
}

/**
 * This function reduces the set of all barcodes down to just the Yeti UPC codes, then adds the inventory values to that data and stores the information on a hidden sheet.
 * 
 * @author Jarren Ralf
 */
function importYetiUPCs()
{
  var item;
  const inventory = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const data = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString())
    .filter(descrip => descrip[2].toUpperCase().includes('YETI')) // Grab only the items that contain YETI in the description
    .map(value => {
      item = inventory.find(sku => sku[6].toUpperCase() === value[1].toUpperCase()) // Match the SKUs in our inventory
      return [value[0], value[3], item[1], '', item[2], item[3], item[4]] // Return the item and the associated inventory values
    })
  
  SpreadsheetApp.getActive().getSheetByName('Yeti UPCs').getRange(2, 1, data.length, 7).setNumberFormat('@').setValues(data)
}

/**
 * This function checks to see if the active sheet is blank. This is done in order to determine if the user needs the Last Imported data to retrieve the header information or not.
 * 
 * @param {Sheet} sheet : The active sheet.
 * @author Jarren Ralf
 */
function isCurrentExportBlank(sheet)
{
  const header = sheet.getSheetValues(1, 2, 2, 4);

  const isOrderBlank = header[0][0] === '' &&
                       header[0][1] === '' &&
                       header[0][2] === 'Type your order description in this cell. (40 characters max)' &&
                       header[0][3] === '' &&
                       header[1][0] === 'Type your order comments in this cell. (75 characters max)' &&
                       header[1][1] === '' &&
                       header[1][2] === '';

  return isOrderBlank;
}

/**
 * This function checks if every value in the import multi-array is blank, which means that the user has
 * highlighted and deleted all of the data.
 * 
 * @param {Object[][]} values : The import data
 * @return {Boolean} Whether the import data is deleted or not
 * @author Jarren Ralf
 */
function isEveryValueBlank(values)
{
  return values.every(arr => arr.every(val => val == '') === true);
}

/**
 * This function checks if the current row of import data is a new purchase order or part of the previous one. 
 * The first order is NOT considered a new order.
 * 
 * @param {Object[][]} data : The import data
 * @param   {Number}    i   : The current row number of the import data
 * @param   {Number}  poNum : The column number that the PO number is found 
 * @return {Boolean} Whether the current row Name parameter matches the previous order (row above the current) Name parameter or not
 * @author Jarren Ralf
 */
function isNewPurchaseOrder(data, i, poNum)
{
  return i !== 1 && data[i][poNum] !== data[i - 1][poNum] && isNotBlank(data[i][poNum]);
}

/**
 * This function checks if a given value is precisely a non-blank string.
 * 
 * @param  {String}  value : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is not-blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(value)
{
  return value !== '';
}

/**
 * This function returns true if the presented number is a UPC-A, false otherwise.
 * 
 * @param {Number} upcNumber : The UPC-A number
 * @returns Whether the given value is a UPC-A or not
 * @author Jarren Ralf
 */
function isUPC_A(upcNumber)
{
  for (var i = 0, sum = 0, upc = upcNumber.toString(); i < upc.length - 1; i++)
    sum += (i % 2 === 0) ? Number(upc[i])*3 : Number(upc[i])

  return upc.endsWith(Math.ceil(sum/10)*10 - sum)
}

/**
 * This function makes the comment lines when a user has clicked on the row numbers and selected insert rows. It adds a 'C' to the first cell
 * and promts the user to type their comment.
 * 
 * @param {Event} e : The event object
 * @throws Throws an error if the script doesn't run
 * @author Jarren Ralf
 */
function makeCommentLines(e)
{
  try
  {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getSheetName();

    if (sheetName === 'Export 1' || sheetName === 'Export 2') // A row is added on the export sheets
    {
      const activeRange = sheet.getActiveRange();
      const row = activeRange.getRow();

      // If a row is added in the middle of the data, then make those rows a 'comment' line (C)
      if (activeRange.getNumColumns() === 0 && row <= sheet.getLastRow() && row >= 2)
        sheet.getRange(activeRange.getRow(), 1, activeRange.getNumRows(), 2)
          .setHorizontalAlignments(new Array(activeRange.getNumRows()).fill(['center', 'left']))
          .setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
          .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
          .setValues(new Array(activeRange.getNumRows()).fill(['C', 'Type your line comments in this cell. (75 characters max)']))
    }
  }
  catch (err)
  {
    var error = err['stack'];
    Logger.log(error);
    Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
    throw new Error(error);
  }
}

/**
 * This function manages the edits that take place on the Export sheets. The user may be appending import data to the sheet,
 * in which case the page needs be reformatted and the data processed. It also handles whether the user is manually changing
 * the vendor or changing a SKU or multiple SKUs.
 * 
 * @param {Event} e : The event object 
 * @param {Number} sheetNumber : The number of the sheet that is being edited, ex. Export '1' vs '2'
 * @throws Throws an error if the script doesn't run
 * @author Jarren Ralf
 */
function manageExportSheetEdit(e, sheetNumber)
{
  var range = e.range;
  var spreadsheet = e.source
  var row = range.rowStart;
  var col = range.columnStart;
  var rowEnd = range.rowEnd;
  var isSingleRow    = row == rowEnd;
  var isSingleColumn = col == range.columnEnd;
  var sheet = SpreadsheetApp.getActiveSheet();
  var NUM_IMPORTS = 1, NUM_COLS = 5, blue = '#e0e9f9';

  try
  {
    if (sheet.getLastColumn() > NUM_COLS) // User may be attempting to Append to the Export sheet (File -> Import -> Append to current sheet)
    {
      if (NUM_IMPORTS === 1) // Only allow 1 export per onChange edit
      {
        NUM_IMPORTS++; // Increment counter so no more exports are attempted
        var currentUsers = []
        currentUsers.push(sheet.getRange('E2').activate().getValue().split(' ', 1)[0]); // Move to the top of the export page

        if (currentUsers[0] === '') // There is no one currently editing this sheet 
        {
          const values = sheet.getRange(row, 1, sheet.getLastRow() - row + 1, sheet.getLastColumn()).getValues();
          spreadsheet.getSheetByName('Last Import ' + sheetNumber).clearContents().getRange(1, 1, values.length, values[0].length).setNumberFormat('@')
            .setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold')
            .setVerticalAlignment('middle').setBackground(null).setValues(values); // Put the imported values on the Last Import sheet
          sheet.getRange(row, 1, rowEnd - row + 1, NUM_COLS).clearContent()
          const vendorName = getVendorName(values)
          Logger.log(vendorName)
          exportInfo(vendorName, values, sheet, false, spreadsheet);
        }
        else // Check the second export sheet because the first one has a current user
        {
          sheetNumber = (sheetNumber === 1) ? sheetNumber++ : sheetNumber--; // Flip the sheet number between 1 and 2 depending on which sheet was originally edited
          var exportSheet = spreadsheet.getSheetByName('Export ' + sheetNumber);
          currentUsers.push(exportSheet.getRange('E2').activate().getValue().split(' ', 1)[0]); // Move to the top of the export page

          if (currentUsers[1] === '') // There is no one currently editing this sheet 
          {
            const values = sheet.getRange(row, 1, sheet.getLastRow() - row + 1, sheet.getLastColumn()).getValues();
            spreadsheet.getSheetByName('Last Import ' + sheetNumber).clearContents().getRange(1, 1, values.length, values[0].length).setNumberFormat('@')
              .setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold')
              .setVerticalAlignment('middle').setBackground(null).setValues(values); // Put the imported values on the Last Import sheet
            sheet.getRange(row, 1, rowEnd - row + 1, NUM_COLS).clearContent()
            const vendorName = getVendorName(values)
            Logger.log(vendorName)
            exportInfo(vendorName, values, exportSheet, false, spreadsheet);
          }
          else // Both the export sheets are in use, therefore, an import like this can not occur. The user must physically go ask someone to stop using, or hit the "done" button.
          {
            sheet.getRange(row, 1, rowEnd - row + 1, NUM_COLS).clearContent()
            Browser.msgBox('You are unable to import because ' + currentUsers[0] + ' and ' + currentUsers[1] + ' are currently using Export 1 and 2, respectively.')
          }
        }
      }

      sheet.deleteColumns(6, sheet.getLastColumn() - range.columnEnd); // Delete all of the extra columns that were a result of appending the data to this sheet
    }
    else if (isSingleRow && isSingleColumn &&  isNotBlank(range.getValue())) // User is editing a single cell AND has not pressed delete
    {    
      // && !userHasPressedDelete(e.value)
      if (row === 1 && col === 5) // User is changing the vendor
      {
        range.setHorizontalAlignment('center')
        const vendorSheet = spreadsheet.getSheetByName('Vendors');
        const vendors = vendorSheet.getSheetValues(2, 1, vendorSheet.getLastRow() - 1, 2);
        const rng = sheet.getRange(1, 3, 1, 3);
        const vals = rng.getValues();

        for (var v = 0; v < vendors.length; v++)
        {
          if (e.value === vendors[v][0])
          { 
            vals[0][0] = vendors[v][1]; // Set the vendor name and number
            vals[0][2] = vendors[v][0];
            sheet.getRange(1, 3, 1, 3).setNumberFormat('@').setHorizontalAlignments([['center', 'left', 'center']]).setFontColor('black').setFontFamily('Arial').setFontLine('none')
              .setFontSizes([[13, 13, 18]]).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(blue).setValues(vals)
            break;
          }
        }
      }
      else if (row > 2 && col === 2) // SKU field is being edited
      {
        const vendorName = sheet.getRange('E1').getValue();

        if (vendorName !== '')
        {
          const vendorSheet = spreadsheet.getSheetByName(vendorName);
          const lastCol = vendorSheet.getLastColumn();
          const data = vendorSheet.getSheetValues(1, 1, vendorSheet.getLastRow(), lastCol);
          
          const newSku = sheet.getRange(row, 2, 1, 3).trimWhitespace().getValues()[0];

          if (newSku[2] != '' && newSku[0] != '') // Make sure the description and sku are not blank
          {
            range.setHorizontalAlignment('left')
            const pntSKU = data[0].indexOf('PNT SKU')
            const pntDescription = data[0].indexOf('PNT Description')

            for (var i = 1; i < data.length; i++)
            {
              if (newSku[2] == data[i][pntDescription]) // Found the item in the Vendors database
              {
                const itemRange = vendorSheet.getRange(i + 1, 1, 1, lastCol);
                const itemValues = itemRange.getValues();
                itemValues[0][pntSKU] = newSku[0];
                itemRange.setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
                  .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(itemValues);
                const description = getAdagioDescription(i + 1, pntSKU + 1, newSku[0], vendorSheet, vendorSheet.getSheetName());
                sheet.getRange(row, 4).setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
                  .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValue(description);
                break;
              }
            }
          }
        }
        else // There is no vendor, and therefore we don't know what data sheet to go and edit
        {
          Browser.msgBox('You must select the Vendor from the dropdown box in cell E1 in order to edit a SKU.')
          sheet.getRange('E1').activate();
        }
      }
    }
    // else if (e.value === undefined && e.oldValue === undefined && isSingleColumn && col == 2 && row > 2) // This is a fill down drag in the SKU column
    // {
    //   range.setHorizontalAlignment('left')
    //   const vendorSheet = spreadsheet.getSheetByName(sheet.getRange('E1').getValue());
    //   const lastCol = vendorSheet.getLastColumn();
    //   const data = vendorSheet.getSheetValues(1, 1, vendorSheet.getLastRow(), lastCol);
    //   const newSkus = sheet.getRange(row, 2, rowEnd - row + 1, 3).trimWhitespace().getValues();
    //   const pntSKU = data[0].indexOf('PNT SKU')
    //   const pntDescription = data[0].indexOf('PNT Description')

    //   for (var j = 0; j < newSkus.length; j++)
    //   {
    //     if (newSkus[j][0] !== '')
    //     {
    //       for (var i = 1; i < data.length; i++)
    //       {
    //         if (newSkus[j][2] == data[i][pntDescription]) // Found the item in the Vendors database
    //         {
    //           var itemRange = vendorSheet.getRange(i + 1, 1, 1, lastCol);
    //           var itemValues = itemRange.getValues();
    //           itemValues[0][pntSKU] = newSkus[j][0];
    //           itemRange.setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
    //             .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(itemValues);
    //           var description = getAdagioDescription(i + 1, pntSKU + 1, newSkus[j][0], vendorSheet, vendorSheet.getSheetName()); 
    //           sheet.getRange(row + j, 4).setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
    //             .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValue(description);
    //           break;
    //         }
    //       }
    //     }
    //   }
    // }
  }
  catch (err)
  {
    if (NUM_IMPORTS === 2) // An import was attempted so delete the appended rows
    {
      sheet.getRange(row, 1, rowEnd - row + 1, NUM_COLS).clearContent()
      sheet.deleteColumns(6, sheet.getLastColumn() - range.columnEnd);
    }

    var error = err['stack'];
    Logger.log(error);
    Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
    throw new Error(error);
  }
}

/**
 * This function checks all of the sheets and determines which one has just been added via File -> Import. It then checks whether
 * there are any Export pages that aren't being used and will repackage the import data into a format that is acceptable for import
 * into Adagio.
 * 
 * @param {Event} e : The event object.
 * @throws Throws an error if the script doesn't run
 * @author Jarren Ralf
 */
function processImportedData(e)
{ 
  try
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, NUM_IMPORTS = 1;

    spreadsheet.toast('Looking for Import...')

    for (var sheet = 0; sheet < sheets.length; sheet++) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      info = [
        sheets[sheet].getLastRow(),
        sheets[sheet].getLastColumn(),
        sheets[sheet].getMaxRows(),
        sheets[sheet].getMaxColumns()
      ]

      // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
      if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || (info[maxRow] === 1000 && info[maxCol] >= 26 && info[numRows] !== 0 && info[numCols] !== 0)) 
      {
        spreadsheet.toast('Possible Import Detected...')

        if (NUM_IMPORTS === 1) // Only allow 1 export per onChange edit
        {
          NUM_IMPORTS++; // Increment counter so no more exports are attempted
          var currentUsers = []
          var exportSheet = sheets[sheets.findIndex(sh => sh.getSheetName() === 'Export 1')]
          currentUsers.push(exportSheet.getRange('E2').activate().getValue().split(' ', 1)[0]); // Move to the top of the export page

          if (currentUsers[0] === '') // A user has not declared that they are using Export 1 (by choosing their name from the drop down in E2)
          {
            const values = sheets[sheet].getRange(1, 1, info[numRows], info[numCols]).getValues();
            sheets[sheets.findIndex(sh => sh.getSheetName() === 'Last Import 1')].clearContents().getRange(1, 1, values.length, values[0].length).setNumberFormat('@')
              .setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold')
              .setVerticalAlignment('middle').setBackground(null).setValues(values);
            var vendorName = getVendorName(values)

            spreadsheet.toast('Importing ' + vendorName + '...')

            exportInfo(vendorName, values, exportSheet, false, spreadsheet, sheets[sheet]);
          }
          else // Check the second export sheet
          {
            exportSheet = sheets[sheets.findIndex(sh => sh.getSheetName() === 'Export 2')]
            currentUsers.push(exportSheet.getRange('E2').activate().getValue().split(' ', 1)[0]); // Move to the top of the export page

            if (currentUsers[1] === '') // A user has not declared that they are using Export 1 (by choosing their name from the drop down in E2)
            {
              const values = sheets[sheet].getRange(1, 1, info[numRows], info[numCols]).getValues();
              sheets[sheets.findIndex(sh => sh.getSheetName() === 'Last Import 2')].clearContents().getRange(1, 1, values.length, values[0].length).setNumberFormat('@')
                .setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold')
                .setVerticalAlignment('middle').setBackground(null).setValues(values);
              var vendorName = getVendorName(values)

              spreadsheet.toast('Importing ' + vendorName + '...')

              exportInfo(vendorName, values, exportSheet, false, spreadsheet, sheets[sheet]);
            }
            else
              Browser.msgBox('You are unable to import because ' + currentUsers[0] + ' and ' + currentUsers[1] + ' are currently using Export 1 and 2, respectively.')
          }
        }

        if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
          spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

        if (vendorName === 'Xtratuf')
        {
          spreadsheet.deleteSheet(spreadsheet.getSheetByName('price list-BG')) // Delete the extra sheets from the Xtratuf excel file
          spreadsheet.deleteSheet(spreadsheet.getSheetByName('Rocky-Durango-Georgia Boot'))
          spreadsheet.deleteSheet(spreadsheet.getSheetByName('Price List')) 
          spreadsheet.deleteSheet(spreadsheet.getSheetByName('price list-OG'))
        }

        break;
      }
    }

    spreadsheet.toast('Data Processing Complete')
  }
  catch (err)
  {
    if (NUM_IMPORTS === 2)
      spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet(s) that was created

    var error = err['stack'];
    Logger.log(error);
    Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
    throw new Error(error);
  }
}

/**
* This function grabs the MAX_NUM_ITEMS most recently created items from the Recent page and displays them on the search page.
*
* @param {Spreadsheet}   spreadsheet   : The active spreadsheet
* @param    {Sheet}    itemSearchSheet : The active sheet
* @author Jarren Ralf
*/
function recentlyCreatedItems(spreadsheet, itemSearchSheet)
{
  const startTime = new Date().getTime();
  const MAX_NUM_ITEMS = 5000;

  if (arguments.length !== 2)
  {
    spreadsheet = SpreadsheetApp.getActive();
    itemSearchSheet = spreadsheet.getActiveSheet();
  }

  const recentData = spreadsheet.getSheetByName('Recent').getSheetValues(2, 1, MAX_NUM_ITEMS, 6);
  itemSearchSheet.getRange(5, 1, MAX_NUM_ITEMS, 6).setNumberFormat('@').setValues(recentData);
  itemSearchSheet.getRange(1, 1, 4, 3).setNumberFormat('@').setValues([[MAX_NUM_ITEMS + " recently created items.", null, null], [null, null, null], 
    [null, null, null], [(new Date().getTime() - startTime)/1000 + ' s', null, 'Order\nQuantity']]);
}

/**
 * This function accepts a purchase order number and attempts to reformat it into the acceptable Adagio standard, which is 'PO0#####'.
 * If the PO number is too long or short, then no changes are made, and it is left to the user to do a manual edit. Since there is no
 * precise standard of operation when it comes to submitting PO numbers, we don't want to delete information that may be important.
 * 
 * @param {String} po : The purchase order number for the current order.
 * @return {String} The purchase order number reformatted into the correct Adagio format.
 * @author Jarren Ralf
 */
function reformatPoNumber(po)
{
  // The length of the PO number is checked so that only numbers that appear to be valid are edited.
  if (po.substring(0, 2) !== 'PO') // The PO number does not start with a 'PO'
    po = (po[0] !== '0' && po.length === 5) ? 'PO0' + po : po = 'PO' + po;
  else if (po[2] !== '0' && po.length === 7) // The PO number starts with 'PO', so check the third character for a '0'
    po = 'PO0' + po.substring(2)

  return po
}

/**
 * This function runs the exportInfo function which takes data from the last import sheet inorder to update missing SKUs, etc.
 * 
 * @author Jarren Ralf
 */
function refresh()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const sheetNames = sheet.getSheetName().split(' ', 2);
  const ui = SpreadsheetApp.getUi()

  if (sheetNames[0] === 'Export')
  {
    if (sheetNames[1] == 1 || sheetNames[1] == 2)
    {
      const values = spreadsheet.getSheetByName('Last Import ' + sheetNames[1]).getDataRange().getValues()
      const vendorName = getVendorName(values)
      const isRefresh = (isCurrentExportBlank(sheet)) ? false : true;

      if (isRefresh)
      {
        const userName = sheet.getSheetValues(2, 5, 1, 1)[0][0].split(' ', 1)[0];
        const response = (isNotBlank(userName)) ? 
            ui.alert('Refresh', 'You are about to refresh the current export data from the last imported packing slip.' + 
            ' Any additional changes that were made since the last import may be lost. Please check with the current user of this sheet, ' + 
            userName + ', before selecting Yes.\n\nDo you want to continue?', ui.ButtonSet.YES_NO) 
          : ui.alert('Refresh', 'You are about to refresh the current export data from the last imported packing slip.' + 
            ' Any additional changes that you have made since the last import may be lost.\n\nDo you want to continue?', ui.ButtonSet.YES_NO)

        if (response === ui.Button.YES)
        {
          const lastRow = sheet.getLastRow();
          const output = exportInfo(vendorName, values, sheet, isRefresh, spreadsheet);

          if (lastRow !== 2)
          {
            const oldOutput = sheet.getSheetValues(3, 1, lastRow - 2, 5);
            const areAllValuesIdentical = oldOutput.every((row, i) => row.every((col, j) => col == output[i][j]) == true);

            if (areAllValuesIdentical)
              ui.alert('No Changes To The Data', 'Please ensure that you added the correct SKUs to the appropriate vendor database and pressed the update button,' + 
                ' located in the top right-hand corner of the database.', ui.ButtonSet.OK)
          }
        }
      }
      else // User clicked refresh while the page is blank, so a regular import is being done
        exportInfo(vendorName, values, sheet, isRefresh, spreadsheet);
    }
    else
      ui.alert('Refresh Error: Current Sheet is Not Supported', 'The refresh function is not supported for ' + sheetNames[0] + ' ' + sheetNames[1] +
      '. Please contact the spreadsheet owner if you want this feature added.', ui.ButtonSet.OK)
  }
  else
    ui.alert('Refresh Error: Wrong Sheet', 'In order to refresh data, please select either the Export 1 or 2 tab below, then try running the function again.', ui.ButtonSet.OK)
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the Inventory page for the items in question.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted.
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited.
 * @param    {Sheet}        sheet    : The sheet that is being edited.
 * @throws Throws an error if the script doesn't run
 * @author Jarren Ralf 
 */
function search(e, sheet)
{
  var spreadsheet = e.source
  var range = e.range;
  var row = range.rowStart;
  var col = range.columnStart;
  var rowEnd = range.rowEnd;
  var colEnd = range.columnEnd;
  var isSingleRow    = row == rowEnd;
  var isSingleColumn = col == colEnd;

  try
  {
    if (isSingleRow)
    {
      if (row === 1 && col === 2 && (isSingleColumn || (rowEnd === 2 && (range.colEnd === 3 || range.colEnd == null))))
      {
        const startTime = new Date().getTime();
        const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
        const functionRunTimeRange = sheet.getRange(3, 1);      // The range that will display the runtimes for the search and formatting
        const itemSearchFullRange = sheet.getRange(5, 1, sheet.getMaxRows() - 4, 6); // The entire range of the Item Search page
        const output = [];
        const searchesOrNot = sheet.getRange(1, 2, 2, 2).clearFormat()                                    // Clear the formatting of the range of the search box
          .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
          .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
          .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
          .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
          .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

        const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

        if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
        {
          spreadsheet.toast('Searching...')

          if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
          {
            const inventorySheet = spreadsheet.getSheetByName('Inventory');
            const data = inventorySheet.getSheetValues(2, 1, inventorySheet.getLastRow() - 1, 6);
            const numSearches = searches.length; // The number searches
            var numSearchWords;

            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      output.push(data[i]);
                      break loop;
                    }
                  }
                  else
                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                }
              }
            }
          }
          else // The word 'not' was found in the search string
          {
            var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

            const inventorySheet = spreadsheet.getSheetByName('Inventory');
            const data = inventorySheet.getSheetValues(2, 1, inventorySheet.getLastRow() - 1, 6);
            const numSearches = searches.length; // The number searches
            var numSearchWords;

            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      for (var l = 0; l < dontIncludeTheseWords.length; l++)
                      {
                        if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                        {
                          if (l === dontIncludeTheseWords.length - 1)
                          {
                            output.push(data[i]);
                            break loop;
                          }
                        }
                        else
                          break;
                      }
                    }
                  }
                  else
                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                }
              }
            }
          }

          const numItems = output.length;

          if (numItems === 0) // No items were found
          {
            sheet.getRange('B1').activate(); // Move the user back to the seachbox
            itemSearchFullRange.clearContent(); // Clear content
            const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
            const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
            searchResultsDisplayRange.setRichTextValue(message);
          }
          else
          {
            sheet.getRange('B5').activate(); // Move the user to the top of the search items
            itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
            sheet.getRange(5, 1, numItems, 6).setNumberFormat('@').setValues(output);
            (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");
          }

          spreadsheet.toast('Searching Complete.')
        }
        else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
        {
          const MAX_NUM_ITEMS = 5000;
          const recentSheet = spreadsheet.getSheetByName('Recent')
          const recentItems = recentSheet.getSheetValues(2, 1, MAX_NUM_ITEMS, 6)
          itemSearchFullRange.clearContent().setBackground('white');
          sheet.getRange(5, 1, MAX_NUM_ITEMS, 6).setNumberFormat('@').setValues(recentItems);
          searchResultsDisplayRange.setValue(MAX_NUM_ITEMS + " recently created items.");
        }
        else
        {
          itemSearchFullRange.clearContent(); // Clear content 
          const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
          const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
          searchResultsDisplayRange.setRichTextValue(message);
        }

        functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
      }
    }
    else if (row > 4) // multiple rows are being edited
    {
      if (isSingleColumn)
      {
        const values = range.getValues().filter(blank => isNotBlank(blank[0]))

        if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
        {
          if (isUPC_A(values[0][0])) // The first value is a UPC-A, so assume all the pasted values are
          {
            const yetiUpcSheet = spreadsheet.getSheetByName('Yeti UPCs')
            const data = spreadsheet.getSheetByName('Yeti UPCs').getSheetValues(2, 1, yetiUpcSheet.getLastRow() - 1, 7)
            var someUpcsNotFound = false, upc;
            
            const upcs = values.map(item => {
            
              for (var i = 0; i < data.length; i++)
                if (item[0] == data[i][0])
                  return data[i].slice(1);

              someUpcsNotFound = true;

              return ['UPC Not Found:', upc, '', '', '', '']
            });

            if (someUpcsNotFound)
            {
              const upcsNotFound = [];
              var isUpcFound;

              const upcsFound = upcs.filter(item => {
                isUpcFound = item[0] !== 'UPC Not Found:'

                if (!isUpcFound)
                  upcsNotFound.push(item)

                return isUpcFound;
              })

              const numUpcsFound = upcsFound.length;
              const numUpcsNotFound = upcsNotFound.length;
              const items = [].concat.apply([], [upcsNotFound, upcsFound]); // Concatenate all of the item values as a 2-D array
              const numItems = items.length
              const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center'])
              const WHITE = new Array(6).fill('white')
              const YELLOW = new Array(6).fill('#ffe599')
              const colours = [].concat.apply([], [new Array(numUpcsNotFound).fill(YELLOW), new Array(numUpcsFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array

              sheet.getRange(5, 1, sheet.getMaxRows() - 4, 6).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
                .offset(0, 0, numItems, 6)
                  .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setBackgrounds(colours).setValues(items)
                .offset(numUpcsNotFound, 0, numUpcsFound, 6).activate()
            }
            else // All UPCs were succefully found
            {
              const numItems = upcs.length
              const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center'])

              sheet.getRange(5, 1, sheet.getMaxRows() - 4, 6).clearContent().setBackground('white').setFontColor('black').offset(0, 0, numItems, 6)
                .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setValues(upcs).activate()
            }
          }
          else // Probably SKU numbers
          {
            const data = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
            var someSKUsNotFound = false, skus;
            
            if (values[0][0].toString().includes('-'))
            {
              skus = values.map(sku => sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).map(item => {
              
                for (var i = 0; i < data.length; i++)
                  if (data[i][6] == item.toString().toUpperCase())
                    return [data[i][0], data[i][1], '',  data[i][2], data[i][3], data[i][4]]

                someSKUsNotFound = true;

                return ['SKU Not Found:', item, '', '', '', '']
              });
            }
            else
            {
              skus = values.map(item => {
              
                for (var i = 0; i < data.length; i++)
                  if (data[i][6] == item[0].toString().toUpperCase())
                    return [data[i][0], data[i][1], '',  data[i][2], data[i][3], data[i][4]]

                someSKUsNotFound = true;

                return ['SKU Not Found:', item[0], '', '', '', '']
              });
            }

            if (someSKUsNotFound)
            {
              const skusNotFound = [];
              var isSkuFound;

              const skusFound = skus.filter(item => {
                isSkuFound = item[0] !== 'SKU Not Found:'

                if (!isSkuFound)
                  skusNotFound.push(item)

                return isSkuFound;
              })

              const numSkusFound = skusFound.length;
              const numSkusNotFound = skusNotFound.length;
              const items = [].concat.apply([], [skusNotFound, skusFound]); // Concatenate all of the item values as a 2-D array
              const numItems = items.length
              const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center'])
              const WHITE = new Array(6).fill('white')
              const YELLOW = new Array(6).fill('#ffe599')
              const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(YELLOW), new Array(numSkusFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array

              sheet.getRange(5, 1, sheet.getMaxRows() - 4, 6).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
                .offset(0, 0, numItems, 6)
                  .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setBackgrounds(colours).setValues(items)
                .offset(numSkusNotFound, 0, numSkusFound, 6).activate()
            }
            else // All SKUs were succefully found
            {
              const numItems = skus.length
              const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center'])

              sheet.getRange(5, 1, sheet.getMaxRows() - 4, 6).clearContent().setBackground('white').setFontColor('black').offset(0, 0, numItems, 6)
                .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setValues(skus).activate()
            }
          }
        }
      }
      else
      {
        const quantityIndex = colEnd - col; // Assume that the quantity is the final column
        const values = range.getValues().filter(blank => isNotBlank(blank[0]))

        /* Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
         * The first value is a UPC-A, so assume all the pasted values are
         */
        if (values.length !== 0 && isUPC_A(values[0][0])) 
        {
          const yetiUpcSheet = spreadsheet.getSheetByName('Yeti UPCs')
          const data = spreadsheet.getSheetByName('Yeti UPCs').getSheetValues(2, 1, yetiUpcSheet.getLastRow() - 1, 7)
          var someUpcsNotFound = false, upc;
          
          const upcs = values.map(item => {
          
            for (var i = 0; i < data.length; i++)
              if (item[0] == data[i][0])
                return [data[i][1], data[i][2], item[quantityIndex], data[i][5], data[i][5], data[i][6]];

            someUpcsNotFound = true;

            return ['UPC Not Found:', item[0], '', '', '', '']
          });

          if (someUpcsNotFound)
          {
            const upcsNotFound = [];
            var isUpcFound;

            const upcsFound = upcs.filter(item => {
              isUpcFound = item[0] !== 'UPC Not Found:'

              if (!isUpcFound)
                upcsNotFound.push(item)

              return isUpcFound;
            })

            const numUpcsFound = upcsFound.length;
            const numUpcsNotFound = upcsNotFound.length;
            const items = [].concat.apply([], [upcsNotFound, upcsFound]); // Concatenate all of the item values as a 2-D array
            const numItems = items.length
            const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center'])
            const WHITE = new Array(6).fill('white')
            const YELLOW = new Array(6).fill('#ffe599')
            const colours = [].concat.apply([], [new Array(numUpcsNotFound).fill(YELLOW), new Array(numUpcsFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array

            sheet.getRange(5, 1, sheet.getMaxRows() - 4, 6).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
              .offset(0, 0, numItems, 6)
                .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setBackgrounds(colours).setValues(items)
              .offset(numUpcsNotFound, 0, numUpcsFound, 6).activate()
          }
          else // All UPCs were succefully found
          {
            const numItems = upcs.length
            const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center'])

            sheet.getRange(5, 1, sheet.getMaxRows() - 4, 6).clearContent().setBackground('white').setFontColor('black').offset(0, 0, numItems, 6)
              .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setValues(upcs).activate()
          }
        }
      }
    }
  }
  catch (err)
  {
    var error = err['stack'];
    Logger.log(error);
    Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
    throw new Error(error);
  }
}

/**
* Sorts data by the categories while ignoring capitals and pushing blanks to the bottom of the list.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCategories(a, b)
{
  return (a[9].toLowerCase() === b[9].toLowerCase()) ? 0 : (a[9] === '') ? 1 : (b[9] === '') ? -1 : (a[9].toLowerCase() < b[9].toLowerCase()) ? -1 : 1;
}

/**
* Sorts data by the created date of the product for the parksville and rupert spreadsheets.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCreatedDate(a, b)
{
  return (a[7] === b[7]) ? 0 : (a[7] < b[7]) ? 1 : -1;
}

/**
 * This function takes the given string and makes sure that each word in the string has a capitalized 
 * first letter followed by lower case.
 * 
 * @param {String} str : The given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function toProper(str)
{
  var numLetters;
  var words = str.toString().split(' ');

  for (var word = 0, string = ''; word < words.length; word++) 
  {
    numLetters = words[word].length;

    if (numLetters == 0) // The "word" is a blank string (a sentence contained 2 spaces)
      continue; // Skip this iterate
    else if (numLetters == 1) // Single character word
    {
      if (words[word][0] !== words[word][0].toUpperCase()) // If the single letter is not capitalized
        words[word] = words[word][0].toUpperCase(); // Then capitalize it
    }
    else
    {
      /* If the first letter is not upper case or the second letter is not lower case, then
       * capitalize the first letter and make the rest of the word lower case.
       */
      if (words[word][0] !== words[word][0].toUpperCase() || words[word][1] !== words[word][1].toLowerCase())
        words[word] = words[word][0].toUpperCase() + words[word].substring(1).toLowerCase();
    }

    string += words[word] + ' '; // Add a blank space at the end
  }

  string = string.slice(0, -1); // Remove the last space

  return string;
}

/**
 * This function brings back the data from the last export sheet.
 * 
 * @author Jarren Ralf
 */
function undone()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const sheetNames = sheet.getSheetName().split(' ', 2);
  const ui = SpreadsheetApp.getUi()

  if (sheetNames[0] === 'Export')
  {
    if (sheetNames[1] == 1 || sheetNames[1] == 2)
    {
      const userName = sheet.getSheetValues(2, 5, 1, 1)[0][0].split(' ', 1)[0];

      if (isNotBlank(userName))
      {
        const response = ui.alert('Undone', 'You are attempting to bring back the last set of exported data (the last time the DONE button was clicked). ' + 
          'If there is any new data on this page then it will be overwritten. Please check with the current user of this sheet, ' + userName +
          ', before selecting Yes.\n\nDo you want to continue?', ui.ButtonSet.YES_NO)

        if (response === ui.Button.YES)
        {
          const values = spreadsheet.getSheetByName('Last Export ' + sheetNames[1]).getDataRange().getValues()
          sheet.getRange(1, 1, values.length, values[0].length).setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial')
            .setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(values)
        }
      }
      else
      {
        const response = ui.alert('Undone', 'You are attempting to bring back the last set of exported data (the last time the DONE button was clicked). ' + 
          'If there is any new data on this page then it will be overwritten.\n\nDo you want to continue?', ui.ButtonSet.YES_NO)

        if (response === ui.Button.YES)
        {
          const values = spreadsheet.getSheetByName('Last Export ' + sheetNames[1]).getDataRange().getValues()
          sheet.getRange(1, 1, values.length, values[0].length).setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial')
            .setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(values)
        }
      }
    }
    else
      ui.alert('Undone Error: Current Sheet is Not Supported', 'The undone function is not supported for ' + sheetNames[0] + ' ' + sheetNames[1] +
      '. Please contact the spreadsheet owner if you want this feature added.', ui.ButtonSet.OK)
  }
  else
    ui.alert('Undone Error: Wrong Sheet', 'In order to recover data from the last export, please select either the Export 1 or 2 tab below, then try running the function again.', ui.ButtonSet.OK)
}

/**
 * This function adds the PNT Description to the vendor data that is the active sheet.
 * 
 * @author Jarren Ralf
 */
function updateVendors_ActiveSheetOnly()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetArray = [spreadsheet.getActiveSheet()];
  updateVendors(spreadsheet, sheetArray);
}

/**
 * This function updates the PNT Description on the particular vendor sheet when the user changes the PNT SKU.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted.
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited.
 * @param    {Sheet}        sheet    : The sheet that is being edited.
 * @throws Throws an error if the script doesn't run
 * @author Jarren Ralf
 */
function updatePntDescription(e, sheet, sheetName)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;
  const rowEnd = range.rowEnd;
  const isSingleRow    = row == rowEnd;
  const isSingleColumn = col == range.columnEnd;
  const header = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn())[0];
  const skuColumn = header.indexOf('PNT SKU') + 1;

  try
  {
    if (isSingleColumn && col === skuColumn && row > 1) // The PNT SKU column is and not the header is being edited
    {
      range.setHorizontalAlignment('left');

      if (isSingleRow)
      {
        const sku = sheet.getRange(row, col, 1, 1).trimWhitespace().getValues()[0][0];

        if (isNotBlank(sku))
          getAdagioDescription(row, col, sku, sheet, sheetName)
      }
      // else if (e.value === undefined && e.oldValue === undefined) // This is a fill down drag
      // {
      //   const pntItemRange = sheet.getRange(row, skuColumn, rowEnd - row + 1, 2).trimWhitespace();
        
      //   getAdagioDescriptions(row, pntItemRange, sheet, sheetName);
      // }
    }
  }
  catch (err)
  {
    var error = err['stack'];
    Logger.log(error);
    Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
    throw new Error(error);
  }
}

/**
 * This function updates the vendor information.
 * 
 * @param {String}       vendorName : The name of the vendor.
 * @param {Number}     vendorNumber : The vendor number which is the id for a particular vendor in the Adagio system.
 * @param {String[][]}      vendors : A list of vendor names and numbers.
 * @param {Range[][]}    vendorsRng : The range of containing the vendor names and numbers.
 * @param {Sheet}       vendorSheet : The sheet that contain the vendor names and numbers.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {Ui}                   ui : An instance of the user interface object.
 * @author Jarren Ralf
 */
function updateVendor(vendorName, vendorNumber, vendors, vendorsRng, vendorSheet, spreadsheet, ui)
{
  vendorNumber = vendorNumber.toString();

  if (vendorNumber.length > 6)
    ui.alert('Vendor Number Too Long', 'The vendor number should only be 6 digits.', ui.ButtonSet.OK)
  else if (vendorNumber.length < 6)
    ui.alert('Vendor Number Too Short', 'The vendor number should be 6 digits.', ui.ButtonSet.OK)
  else if (isNaN(vendorNumber))
    ui.alert('Invalid Vendor Number', 'The vendor number should contain only numerals.', ui.ButtonSet.OK)
  else
  {
    const Name = 0, Number = 1;
    var isNewVendor = true; // By default a new vendor is being added, unless the name or number already matches one on record

    for (var v = 0; v < vendors.length; v++)
    {
      if (vendors[v][Name].toString().toLowerCase() === vendorName.toString().toLowerCase() && vendors[v][Number] === vendorNumber) // Both name and number pair were found
      {
        ui.alert('Vendor Found', vendorName + ' (' + vendorNumber + ') is already in the vendor list.', ui.ButtonSet.OK)
        isNewVendor = false;
        break;
      }
      else if (vendors[v][Name].toString().toLowerCase() === vendorName.toString().toLowerCase()) // Vendor name was found in the list
      {
        const response = ui.alert('Update Vendor', vendorName + ' is saved with vendor # ' + vendors[v][Number] + '.\n\n' + 
          'Would you like to change it to ' + vendorNumber + '?', Browser.Buttons.YES_NO)

        if (response === ui.Button.YES)
        {
          vendors[v][Number] = vendorNumber;
          vendorsRng.setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
            .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(vendors);
          spreadsheet.toast(vendorName + ' vendor # was changed to ' + vendorNumber + '.', 'Vendor Updated')
        }
        isNewVendor = false;
        break;
      }
      else if (vendors[v][Number] === vendorNumber) // Vendor number was found in the list
      {
        const response = ui.alert('Update Vendor', 'Vendor # ' + vendorNumber + ' was saved with vendor name ' + vendors[v][Name] + '.\n\n' + 
          'Would you like to change the name to ' + vendorName + '?', Browser.Buttons.YES_NO)

        if (response === ui.Button.YES)
        {
          vendors[v][Name] = vendorName;
          vendorsRng.setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
            .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(vendors);
          spreadsheet.toast(' Vendor # ' + vendorNumber + ' is now associated with ' + vendorName + '.', 'Vendor Updated')
        }
        isNewVendor = false;
        break;
      }
    }

    if (isNewVendor)
    {
      vendors.push([vendorName, vendorNumber])
      vendors.sort();
      vendorSheet.getRange(2, 1, vendors.length, 2).setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none')
        .setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(vendors)
      spreadsheet.toast(vendorName + ' (' + vendorNumber + ') has been added to the vendor list.', 'New Vendor Added')
    }
  }
}

/**
 * This function adds the PNT Description to the vendors data and subsequently updates it.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param   {Sheet[]}      sheets   : The active sheet in an array.
 * @author Jarren Ralf
 */
function updateVendors(spreadsheet, sheets)
{
  if (arguments.length !== 2)
  {
    spreadsheet = SpreadsheetApp.getActive();
    sheets = spreadsheet.getSheets();
  }

  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const sheetNames = sheets.map(sheet => sheet.getSheetName());
  var range = new Array(sheets.length), data = new Array(sheets.length);

  for (var s = 0; s < sheets.length; s++)
  {
    if (sheetNames[s] === 'Grundens')
    {
      range[s] = sheets[s].getDataRange();
       data[s] = range[s].getValues();
      const sku           = data[s][0].indexOf("PNT SKU")
      const description   = data[s][0].indexOf("PNT Description")
      const grunDescrip   = data[s][0].indexOf("Name")
      const grunColour    = data[s][0].indexOf("Color")
      const grunSize      = data[s][0].indexOf("Size 1")
      const grunSubCat    = data[s][0].indexOf("Subcategory")
      const grunCat       = data[s][0].indexOf("Category")
      const grunPrice     = data[s][0].indexOf("Price 1 USD")
      const grunStyleNum  = data[s][0].indexOf("Style Number")
      const grunColorCode = data[s][0].indexOf("Color Code")
      var proposedNewSKU;

      for (var i = 1; i < data[s].length; i++)
      {
        if (data[s][i][sku] === '' || data[s][i][sku].toString().split(' - ', 1)[0] === 'SKU_NOT_FOUND' || data[s][i][sku].toString().split(' - ', 1)[0] === 'NEW_ITEM_ADDED')
        {
          if (data[s][i][sku] === '') data[s][i][sku] = "SKU_NOT_FOUND";

          proposedNewSKU = data[s][i][grunStyleNum].toString() + ((data[s][i][grunColorCode].toString().length == 1) ? '00' + data[s][i][grunColorCode].toString() : 
                                                              (data[s][i][grunColorCode].toString().length == 2) ?  '0' + data[s][i][grunColorCode].toString() : 
                                                                data[s][i][grunColorCode].toString()) + ((data[s][i][grunSize] === 'XXL') ? '2XL' : data[s][i][grunSize])
          data[s][i][description] = data[s][i][grunDescrip] + ' - ' + data[s][i][grunColour] + ' - '       + data[s][i][grunSize] + ' - ' + 
                                  ((data[s][i][grunSubCat] !== '') ? data[s][i][grunSubCat] + ' - ' : '- ') +
                                  ((data[s][i][grunCat] !== '') ? data[s][i][grunCat] + ' - ' : '- ') + 'Cost: $' + Number(data[s][i][grunPrice]).toFixed(2) + ' - ' + proposedNewSKU;
        }
        else
        {
          for (var j = 0; j < csvData.length; j++)
          {
            if (data[s][i][sku].toString().split(' - ', 1)[0] == csvData[j][6]) // Match the SKUs
            {
              data[s][i][description] = csvData[j][1] // Add the adagio description
              data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] // Remove " - create new?" from the SKU
              break;
            }
          }

          if (j === csvData.length) // SKU was not found in the Adagio data
          {
            data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] + " - create new?";
            proposedNewSKU = data[s][i][grunStyleNum].toString() + ((data[s][i][grunColorCode].toString().length == 1) ? '00' + data[s][i][grunColorCode].toString() : 
                                                                  (data[s][i][grunColorCode].toString().length == 2) ?  '0' + data[s][i][grunColorCode].toString() : 
                                                                  data[s][i][grunColorCode].toString()) + ((data[s][i][grunSize] === 'XXL') ? '2XL' : data[s][i][grunSize])
            data[s][i][description] = "SKU not in Adagio. " + data[s][i][grunDescrip] + ' - ' + data[s][i][grunColour] + ' - '       + data[s][i][grunSize] + ' - ' + 
                                    ((data[s][i][grunSubCat] !== '') ? data[s][i][grunSubCat] + ' - ' : '- ') +
                                    ((data[s][i][grunCat] !== '') ? data[s][i][grunCat] + ' - ' : '- ') + 'Cost: $' + Number(data[s][i][grunPrice]).toFixed(2) + ' - ' + proposedNewSKU;
          }
        }
      }

      range[s].setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
        .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(data[s])
    }
    else if (sheetNames[s] === 'Helly Hansen')
    {
      range[s] = sheets[s].getDataRange();
       data[s] = range[s].getValues();
      const sku              = data[s][0].indexOf("PNT SKU")
      const description      = data[s][0].indexOf("PNT Description")
      const hellyDescription = data[s][0].indexOf("Style Name")
      const hellyColour      = data[s][0].indexOf("Color Name")
      const hellySize        = data[s][0].indexOf("Size")
      const hellyPrice       = data[s][0].indexOf("Wholesale Price")
      const hellySKU         = data[s][0].indexOf("SKU")
      var proposedNewSKU;

      for (var i = 1; i < data[s].length; i++)
      {
        if (data[s][i][sku] === '' || data[s][i][sku].toString().split(' - ', 1)[0] === 'SKU_NOT_FOUND' || data[s][i][sku].toString().split(' - ', 1)[0] === 'NEW_ITEM_ADDED')
        {
          if (data[s][i][sku] === '') data[s][i][sku] = "SKU_NOT_FOUND";

          proposedNewSKU = data[s][i][hellySKU].replace(/_|-/g,'')
          data[s][i][description] = data[s][i][hellyDescription] + ' - ' + data[s][i][hellyColour] + ' - ' + data[s][i][hellySize] + 
                                ' - Cost: $' + Number(data[s][i][hellyPrice]).toFixed(2) + ' - ' + proposedNewSKU;
        }
        else
        {
          for (var j = 0; j < csvData.length; j++)
          {
            if (data[s][i][sku].toString().split(' - ', 1)[0] == csvData[j][6]) // Match the SKUs
            {
              data[s][i][description] = csvData[j][1] // Add the adagio description
              data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] // Remove " - create new?" from the SKU
              break;
            }
          }

          if (j === csvData.length) // SKU was not found in the Adagio data
          {
            data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] + " - create new?";
            proposedNewSKU = data[s][i][hellySKU].replace(/_|-/g,'')
            data[s][i][description] = "SKU not in Adagio. " + data[s][i][hellyDescription] + ' - ' + data[s][i][hellyColour] 
              + ' - ' + data[s][i][hellySize] + ' - Cost: $' + Number(data[s][i][hellyPrice]).toFixed(2) + ' - ' + proposedNewSKU
          }
        }
      }

      range[s].setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
        .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(data[s])
    }
    else if (sheetNames[s] === 'Xtratuf')
    {
      range[s] = sheets[s].getDataRange();
       data[s] = range[s].getValues();
      const sku                = data[s][0].indexOf("PNT SKU")
      const description        = data[s][0].indexOf("PNT Description")
      const xtratufDescription = data[s][0].indexOf("Name/Nom/Description")
      const xtratufCategory    = data[s][0].indexOf("Category / Catégorie")
      const xtratufColour      = data[s][0].indexOf("Color / Couleur")
      const xtratufSize        = data[s][0].indexOf("Sizes / Tailles")
      const xtratufPrice       = data[s][0].indexOf("Purchase Price/\nPrix d\’achat ")
      const xtratufSku         = data[s][0].indexOf("Stock# / Nº de nomenclature")
      var proposedNewSKU;

      for (var i = 1; i < data[s].length; i++)
      {
        if (data[s][i][sku] === '' || data[s][i][sku].toString().split(' - ', 1)[0] === 'SKU_NOT_FOUND' || data[s][i][sku].toString().split(' - ', 1)[0] === 'NEW_ITEM_ADDED')
        {
          if (data[s][i][sku] === '') data[s][i][sku] = "SKU_NOT_FOUND";

          proposedNewSKU = getProposedNewXtratufSKU(data[s][i][xtratufSku], data[s][i][xtratufSize]);
          data[s][i][description] = data[s][i][xtratufDescription] + ' - ' + data[s][i][xtratufCategory] + ' - ' + data[s][i][xtratufColour] + 
                                ' - ' + data[s][i][xtratufSize] + ' - Cost: $' + Number(data[s][i][xtratufPrice]).toFixed(2) + ' - ' + proposedNewSKU
        }
        else
        {
          for (var j = 0; j < csvData.length; j++)
          {
            if (data[s][i][sku].toString().split(' - ', 1)[0] == csvData[j][6]) // Match the SKUs
            {
              data[s][i][description] = csvData[j][1] // Add the adagio description
              data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] // Remove " - create new?" from the SKU
              break;
            }
          }

          if (j === csvData.length) // SKU was not found in the Adagio data
          {
            data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] + " - create new?";
            proposedNewSKU = getProposedNewXtratufSKU(data[s][i][xtratufSku], data[s][i][xtratufSize]);
            data[s][i][description] = "SKU not in Adagio. " + data[s][i][xtratufDescription] + ' - ' + data[s][i][xtratufCategory] + ' - ' + data[s][i][xtratufColour] + 
                                  ' - ' + data[s][i][xtratufSize] + ' - Cost: $' + Number(data[s][i][xtratufPrice]).toFixed(2) + ' - ' + proposedNewSKU
          }
        }
      }

      range[s].setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
        .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(data[s])
    }
    else if (sheetNames[s] === 'Yeti')
    {
      range[s] = sheets[s].getDataRange();
       data[s] = range[s].getValues();
      const sku             = data[s][0].indexOf("PNT SKU")
      const description     = data[s][0].indexOf("PNT Description")
      const yetiDescription = data[s][0].indexOf("DESCRIPTION")
      const yetiCategory    = data[s][0].indexOf("CATEGORY")
      const yetiPrice       = data[s][0].indexOf("DEALER PRICE")
      var proposedNewSKU;

      for (var i = 1; i < data[s].length; i++)
      {
        if (data[s][i][sku] === '' || data[s][i][sku].toString().split(' - ', 1)[0] === 'SKU_NOT_FOUND' || data[s][i][sku].toString().split(' - ', 1)[0] === 'NEW_ITEM_ADDED')
        {
          if (data[s][i][sku] === '') data[s][i][sku] = "SKU_NOT_FOUND";

          proposedNewSKU = getProposedNewYetiSKU(data[s][i][yetiCategory], data[s][i][yetiDescription]);
          data[s][i][description] = data[s][i][yetiDescription] + ' - ' + data[s][i][yetiCategory] + ' - Cost: $' + Number(data[s][i][yetiPrice]).toFixed(2) + ' - ' + proposedNewSKU;
        }
        else
        {
          for (var j = 0; j < csvData.length; j++)
          {
            if (data[s][i][sku].toString().split(' - ', 1)[0] == csvData[j][6]) // Match the SKUs
            {
              data[s][i][description] = csvData[j][1] // Add the adagio description
              data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] // Remove " - create new?" from the SKU
              break;
            }
          }

          if (j === csvData.length) // SKU was not found in the Adagio data
          {
            data[s][i][sku] = data[s][i][sku].toString().split(' - ', 1)[0] + " - create new?";
            proposedNewSKU = getProposedNewYetiSKU(data[s][i][yetiCategory], data[s][i][yetiDescription]);
            data[s][i][description] = "SKU not in Adagio. " + data[s][i][yetiDescription] + ' - ' + data[s][i][yetiCategory] + 
                                  ' - Cost: $' + Number(data[s][i][yetiPrice]).toFixed(2) + ' - ' + proposedNewSKU
          }
        }
      }

      range[s].setNumberFormat('@').setHorizontalAlignment('left').setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10)
        .setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(data[s])
    }
  }
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning true if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is undefined or not.
* @author Jarren Ralf
*/
function userHasPressedDelete(value)
{
  return value === undefined;
}

/**
 * This function creates the export data for an Xtratuf purchase order.
 * 
 * @param {Object[][]}       poData : The purchase order data that was just uploaded.
 * @param {String}          Xtratuf : The name of the vendor, in this case Xtratuf.
 * @param {Sheet}       exportSheet : The sheet that the data will be exported to.
 * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @return {Object[][]}      output : The data created for the export sheet.
 * @author Jarren Ralf
 */
function xtratuf(poData, Xtratuf, exportSheet, isRefresh, spreadsheet)
{
  const xtratufSheet = spreadsheet.getSheetByName('Xtratuf');
  const xtratufData = xtratufSheet.getDataRange().getValues();
  const poRowNum = 4, headerRowNum = 14;
  var addToXtratufData = [], proposedNewSKU, poSizeTypeRowNum, size_FromHeader;

  // Xtratuf Data
  const stockNum       = xtratufData[0].indexOf("Stock# / Nº de nomenclature"); 
  const size           = xtratufData[0].indexOf("Sizes / Tailles"); 
  const cost           = xtratufData[0].indexOf("Purchase Price/\nPrix d\’achat ");
  const pntSKU         = xtratufData[0].indexOf("PNT SKU");
  const pntDescription = xtratufData[0].indexOf("PNT Description");

  // Purchase Order Data
  const sku         = poData[headerRowNum].indexOf("Stock# / Nº de nomenclature"); 
  const brand       = poData[headerRowNum].indexOf("Brand/\nMarque"); 
  const category    = poData[headerRowNum].indexOf("Category / Catégorie");
  const description = poData[headerRowNum].indexOf("Name/Nom/Description"); 
  const colour      = poData[headerRowNum].indexOf('Color / Couleur')
  const price       = poData[headerRowNum].indexOf("Purchase Price/\nPrix d\’achat "); 
  const retailPrice = poData[headerRowNum].indexOf("MSRP/\nPDSF $  CAD"); 
  const poNum       = poData[poRowNum].indexOf("Purchase Order / Bon de commande");
  const po = reformatPoNumber(poData[poRowNum + 1][poNum].toString());
  const vendorNumber = getVendorNumber(Xtratuf, spreadsheet);
  var currentEditor = exportSheet.getSheetValues(2, 5, 1, 1)[0][0];

  if (currentEditor === '')
    currentEditor = 'Someone is currently editing this PO'

  const output = (isRefresh) ? [] : [['H', po, vendorNumber, 'Type your order description in this cell. (40 characters max)', Xtratuf], // Header line
                                     ['C', 'Type your order comments in this cell. (75 characters max)', null, null, currentEditor]] // Comment Line
  
  for (var i = headerRowNum + 4; i < poData.length; i++)
  {
    for (var j = 1; j < xtratufData.length; j++)
    {
      // Find the item in the Xtratuf data base
      if (poData[i][sku] == xtratufData[j][stockNum])
      {
        switch (poData[i][category]) // The category determines which row of size types is used
        {
          case "Men's / Homme":
          case "Women's / Femme":
          case "Unisex":

            poSizeTypeRowNum = headerRowNum + 3;

            for (var k = 6; k < 29; k++)
            {
              if (isNotBlank(poData[i][k]))
              {
                size_FromHeader = poData[poSizeTypeRowNum][k];

                while (poData[i][sku] == xtratufData[j][stockNum])
                {
                  if (size_FromHeader == xtratufData[j][size])
                  {
                    output.push(['R', xtratufData[j][pntSKU], poData[i][k], xtratufData[j][pntDescription], Number(xtratufData[j][cost]).toFixed(2)]) // Receiving Line
                    break;
                  }

                  j++;
                }
              }
            }

            break;
          case "Kids / Enfants":

            poSizeTypeRowNum = headerRowNum + 2;

            for (var k = 6; k < 24; k++)
            {
              if (isNotBlank(poData[i][k]))
              {
                size_FromHeader = poData[poSizeTypeRowNum][k];

                while (poData[i][sku] == xtratufData[j][stockNum])
                {
                  if (size_FromHeader == xtratufData[j][size])
                  {
                    output.push(['R', xtratufData[j][pntSKU], poData[i][k], xtratufData[j][pntDescription], Number(xtratufData[j][cost]).toFixed(2)]) // Receiving Line
                    break;
                  }
                  
                  j++;
                }
              }
            }

            break;
          case "APP/ACC":

            poSizeTypeRowNum = headerRowNum + 1;

            for (var k = 6; k < 15; k++)
            {
              if (isNotBlank(poData[i][k]))
              {
                if (k === 6) // One Size
                {
                  output.push(['R', xtratufData[j][pntSKU], poData[i][k], xtratufData[j][pntDescription], Number(xtratufData[j][cost]).toFixed(2)]) // Receiving Line
                  break;
                }
                else
                {
                  size_FromHeader = poData[poSizeTypeRowNum][k];

                  while (poData[i][sku] == xtratufData[j][stockNum])
                  {
                    if (size_FromHeader == xtratufData[j][size])
                    {
                      output.push(['R', xtratufData[j][pntSKU], poData[i][k], xtratufData[j][pntDescription], Number(xtratufData[j][cost]).toFixed(2)]) // Receiving Line
                      break;
                    }
                    
                    j++;
                  }
                }
              }
            }

            break;
        }

        break;
      }
    }

    if (j === xtratufData.length) // The item(s) that were ordered were not found on the Xtratuf data sheet
    {
      switch (poData[i][category]) // The category determines which row of size types is used
      {
        case "Men's / Homme":
        case "Women's / Femme":
        case "Unisex":

          poSizeTypeRowNum = headerRowNum + 3;

          for (var k = 6; k < 29; k++)
          {
            if (isNotBlank(poData[i][k]))
            {
              proposedNewSKU = getProposedNewXtratufSKU(poData[i][stockNum], poData[i][size]);

              addToXtratufData.push([poData[i][sku], poData[i][brand], poData[i][category], poData[i][description], poData[i][colour], poData[poSizeTypeRowNum][k], poData[i][price], 
                poData[i][retailPrice], 'NEW_ITEM_ADDED', poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[poSizeTypeRowNum][k] + ' - Cost: $' 
                + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
              output.push(['R', 'NEW_ITEM_ADDED', poData[i][k], poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[poSizeTypeRowNum][k] 
                + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)]) // Receiving Line
            }
          }

          break;
        case "Kids / Enfants":

          poSizeTypeRowNum = headerRowNum + 2;

          for (var k = 6; k < 24; k++)
          {
            if (isNotBlank(poData[i][k]))
            {
              proposedNewSKU = getProposedNewXtratufSKU(poData[i][stockNum], poData[i][size]);

              addToXtratufData.push([poData[i][sku], poData[i][brand], poData[i][category], poData[i][description], poData[i][colour], poData[poSizeTypeRowNum][k], poData[i][price], 
                poData[i][retailPrice], 'NEW_ITEM_ADDED', poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[poSizeTypeRowNum][k] + ' - Cost: $' 
                + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
              output.push(['R', 'NEW_ITEM_ADDED', poData[i][k], poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[poSizeTypeRowNum][k] 
                + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)]) // Receiving Line
            }
          }

          break;
        case "APP/ACC":

          poSizeTypeRowNum = headerRowNum + 1;

          for (var k = 6; k < 15; k++)
          {
            if (isNotBlank(poData[i][k]))
            {
              if (k === 6) // One Size 
              {
                proposedNewSKU = getProposedNewXtratufSKU(poData[i][stockNum], poData[i][size]);

                addToXtratufData.push([poData[i][sku], poData[i][brand], poData[i][category], poData[i][description], poData[i][colour], poData[i][size], poData[i][price], 
                  poData[i][retailPrice], 'NEW_ITEM_ADDED', poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[i][size] + ' - Cost: $' 
                  + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
                output.push(['R', 'NEW_ITEM_ADDED', poData[i][k], poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[i][size] 
                  + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)]) // Receiving Line            
              }
              else
              {
                proposedNewSKU = getProposedNewXtratufSKU(poData[i][stockNum], poData[i][size]);
 
                addToXtratufData.push([poData[i][sku], poData[i][brand], poData[i][category], poData[i][description], poData[i][colour], poData[poSizeTypeRowNum][k], poData[i][price], 
                  poData[i][retailPrice], 'NEW_ITEM_ADDED', poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[poSizeTypeRowNum][k] + ' - Cost: $' 
                  + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU])
                output.push(['R', 'NEW_ITEM_ADDED', poData[i][k], poData[i][description] + ' - ' + poData[i][category] + ' - ' + poData[i][colour] + ' - ' + poData[poSizeTypeRowNum][k] 
                  + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU, Number(poData[i][price]).toFixed(2)]) // Receiving Line
              }
            }
          }

          break;
      }
    }
  }

  if (addToXtratufData.length !== 0) // Add items to the Xtratuf data that were on the PO but not in the Xtratuf data
  {
    xtratufSheet.showSheet().getRange(xtratufSheet.getLastRow() + 1, 1, addToXtratufData.length, addToXtratufData[0].length).activate().setHorizontalAlignment('left')
      .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
      .setValues(addToXtratufData)
    spreadsheet.toast('Add PNT SKUs to the Xtratuf data.', '⚠️ New Xtratuf Items ⚠️', 30)
  }

  return output;
}

// /**
//  * This function creates the export data for a Yeti purchase order.
//  * 
//  * @param {Object[][]}       poData : The purchase order data that was just uploaded.
//  * @param {String}             Yeti : The name of the vendor, in this case Yeti.
//  * @param {Sheet}       exportSheet : The sheet that the data will be exported to.
//  * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
//  * @param {Spreadsheet} spreadsheet : The active spreadsheet.
//  * @return {Object[][]}      output : The data created for the export sheet.
//  * @author Jarren Ralf
//  */
// function yeti(poData, Yeti, exportSheet, isRefresh, spreadsheet)
// {
//   const yetiSheet = spreadsheet.getSheetByName('Yeti');
//   const yetiData = yetiSheet.getDataRange().getValues();
//   const vendorNumber = getVendorNumber(Yeti, spreadsheet);
//   var output = [], addToYetiData = [], poRowNum = 2, headerRowNum = 24, proposedNewSKU;

//   // Yeti Data
//   const cost           = yetiData[0].indexOf("DEALER PRICE");
//   const partNum        = yetiData[0].indexOf("NEW YETI PART#");
//   const pntSKU         = yetiData[0].indexOf("PNT SKU");
//   const pntDescription = yetiData[0].indexOf("PNT Description");

//   // Purchase Order Data
//   const qty         = poData[headerRowNum].indexOf("ORDER QTY"); 
//   const sku         = poData[headerRowNum].indexOf("NEW YETI PART#"); 
//   const description = poData[headerRowNum].indexOf("DESCRIPTION"); 
//   const category    = poData[headerRowNum].indexOf("CATEGORY");
//   const price       = poData[headerRowNum].indexOf("DEALER PRICE"); 
//   const poNum       = poData[poRowNum].indexOf("P.O #") + 1;

//   for (var i = headerRowNum + 1; i < poData.length; i++)
//   {
//     if (poData[i][qty] !== '' && poData[i][qty] !== 'ORDER QTY') // Skip the items that we didn't order
//     {
//       for (var j = 1; j < yetiData.length; j++)
//       {
//         // Find the item in the Yeti data base
//         if (poData[i][sku] == yetiData[j][partNum])
//         {
//           output.push(['R', yetiData[j][pntSKU], poData[i][qty], yetiData[j][pntDescription], Number(yetiData[j][cost]).toFixed(2)]) // Receiving Line
//           break;
//         }
//       }

//       if (j === yetiData.length) // The item(s) that were ordered were not found on the Yeti data sheet
//       {
//         proposedNewSKU = getProposedNewYetiSKU(poData[i][category], poData[i][description]);
//         addToYetiData.push(poData[i].push('NEW_ITEM_ADDED', poData[i][description] + ' - ' + poData[i][category] + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU))
//         output.push(['R', 'NEW_ITEM_ADDED', poData[i][qty], poData[i][description] + ' - ' + poData[i][category] + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU,
//                       Number(poData[i][price]).toFixed(2)]) // Receiving Line
//       }
//     }
//   }

//   if (addToYetiData.length !== 0) // Add items to the Yeti data that were on the PO but not in the Yeti data
//   {
//     yetiSheet.showSheet().getRange(yetiSheet.getLastRow() + 1, 1, addToYetiData.length, addToYetiData[0].length).activate().setHorizontalAlignment('left')
//       .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
//       .setValues(addToYetiData)
//     spreadsheet.toast('Add PNT SKUs to the Yeti data.', '⚠️ New Yeti Items ⚠️', 30)
//   }

//   if (output.length !== 0 && !isRefresh)
//   {
//     var currentEditor = exportSheet.getSheetValues(2, 5, 1, 1)[0][0];
//     var po = reformatPoNumber(poData[poRowNum][poNum].toString());

//     if (currentEditor === '')
//       currentEditor = 'Someone is currently editing this PO'

//     output.unshift(
//       ['H', po, vendorNumber, 'Type your order description in this cell. (40 characters max)', Yeti], // Header line
//       ['C', 'Type your order comments in this cell. (75 characters max)', null, null, currentEditor] // Comment Line
//     )
//   }
//   return output;
// }

/**
 * This function creates the export data for a Yeti purchase order.
 * 
 * @param {Object[][]}       poData : The purchase order data that was just uploaded.
 * @param {String}             Yeti : The name of the vendor, in this case Yeti.
 * @param {Sheet}       exportSheet : The sheet that the data will be exported to.
 * @param {Boolean}       isRefresh : A boolean representing whether the user has clicked refresh or not.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @return {Object[][]}      output : The data created for the export sheet.
 * @author Jarren Ralf
 */
function yeti(poData, Yeti, exportSheet, isRefresh, spreadsheet)
{
  const yetiSheet = spreadsheet.getSheetByName('Yeti');
  const yetiData = yetiSheet.getDataRange().getValues();
  const poRowNum = 2, headerRowNum = 24;
  var addToYetiData = [], proposedNewSKU;

  // Yeti Data
  const cost           = yetiData[0].indexOf("DEALER PRICE");
  const partNum        = yetiData[0].indexOf("NEW YETI PART#");
  const pntSKU         = yetiData[0].indexOf("PNT SKU");
  const pntDescription = yetiData[0].indexOf("PNT Description");

  // Purchase Order Data
  const qty         = poData[headerRowNum].indexOf("ORDER QTY"); 
  const sku         = poData[headerRowNum].indexOf("NEW YETI PART#"); 
  const description = poData[headerRowNum].indexOf("DESCRIPTION"); 
  const category    = poData[headerRowNum].indexOf("CATEGORY");
  const price       = poData[headerRowNum].indexOf("DEALER PRICE"); 
  const poNum       = poData[poRowNum].indexOf("P.O #") + 1;
  const po = reformatPoNumber(poData[poRowNum][poNum].toString());
  const vendorNumber = getVendorNumber(Yeti, spreadsheet);
  var currentEditor = exportSheet.getSheetValues(2, 5, 1, 1)[0][0];
    
  if (currentEditor === '')
    currentEditor = 'Someone is currently editing this PO'

  const output = (isRefresh) ? [] : [['H', po, vendorNumber, 'Type your order description in this cell. (40 characters max)', Yeti], // Header line
                                     ['C', 'Type your order comments in this cell. (75 characters max)', null, null, currentEditor]] // Comment Line
  
  for (var i = headerRowNum + 1; i < poData.length; i++)
  {
    if (poData[i][qty] !== '' && poData[i][qty] !== 'ORDER QTY') // Skip the items that we didn't order
    {
      for (var j = 1; j < yetiData.length; j++)
      {
        // Find the item in the Yeti data base
        if (poData[i][sku] == yetiData[j][partNum])
        {
          output.push(['R', yetiData[j][pntSKU], poData[i][qty], yetiData[j][pntDescription], Number(yetiData[j][cost]).toFixed(2)]) // Receiving Line
          break;
        }
      }

      if (j === yetiData.length) // The item(s) that were ordered were not found on the Yeti data sheet
      {
        proposedNewSKU = getProposedNewYetiSKU(poData[i][category], poData[i][description]);
        addToYetiData.push(poData[i].push('NEW_ITEM_ADDED', poData[i][description] + ' - ' + poData[i][category] + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU))
        output.push(['R', 'NEW_ITEM_ADDED', poData[i][qty], poData[i][description] + ' - ' + poData[i][category] + ' - Cost: $' + Number(poData[i][price]).toFixed(2) + ' - ' + proposedNewSKU,
                      Number(poData[i][price]).toFixed(2)]) // Receiving Line
      }
    }
  }

  if (addToYetiData.length !== 0) // Add items to the Yeti data that were on the PO but not in the Yeti data
  {
    yetiSheet.showSheet().getRange(yetiSheet.getLastRow() + 1, 1, addToYetiData.length, addToYetiData[0].length).activate().setHorizontalAlignment('left')
      .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null)
      .setValues(addToYetiData)
    spreadsheet.toast('Add PNT SKUs to the Yeti data.', '⚠️ New Yeti Items ⚠️', 30)
  }

  return output;
}
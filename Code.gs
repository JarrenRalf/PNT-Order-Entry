/**
 * This function handles the on edit events in this spreadsheet pertaining to the Item Search sheet only (all other sheets will be protected).
 * This function is looking for the user searching for items and it is making appropriate changes to the data when a user deletes items from their order.
 * 
 * @param {Event Object} e : The event object
 */
function onEdit(e)
{ 
  const range = e.range;
  const col = range.columnStart;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const isSingleRow = row == rowEnd;
  const isSingleColumn = col == range.columnEnd;
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();

  if (isSingleColumn)
  {
    if (row == 1 && col == 1 && (rowEnd == null || rowEnd == 2 || isSingleRow)) // Item Search
      search(e, spreadsheet, sheet);
    else if (row == 1 && col == 9 && (rowEnd == null || isSingleRow)) // Customer Search
      search_Customer(spreadsheet, sheet);
    else if (row == 2 && col == 2) // Customer Selection
      customerSelection(range, spreadsheet);
    else if (row > 4 && col == 8) // Possibly deleting items from order
      deleteItemsFromOrder(sheet, range, range.getValue(), row, isSingleRow, isSingleColumn, spreadsheet);
  }
}

/**
 * This function processes the imported data.
 * 
 * @param {Event Object} e : The event object from an installed onChange trigger.
 */
function onChange(e)
{
  try
  {
    processImportedData(e)
  }
  catch (error)
  {
    Logger.log(error['stack'])
    Browser.msgBox(error['stack'])
  }
}

/**
 * This function creates a menu item.
 */
function onOpen()
{
  SpreadsheetApp.getUi().createMenu('Export').addItem('Complete', 'completeOrder').addToUi();
}

/**
 * This function identifies all of the cells that the user has selected and moves those items to the order portion of the Item Search sheet.
 * 
 * @author Jarren Ralf
 */
function addSelectedItemsToOrder()
{
  const startTime = new Date().getTime(); // Used for the function runtime
  var firstRows = [], firstCols = [], lastRows = [], lastCols = [], itemValues = [], splitDescription, sku, uom;
  const sheet = SpreadsheetApp.getActiveSheet();

  sheet.getActiveRangeList().getRanges().map((rng, r) => {
    firstRows.push(rng.getRow());
    lastRows.push(rng.getLastRow());
    firstCols.push(rng.getColumn());
    lastCols.push(rng.getLastColumn());
    itemValues.push(...sheet.getSheetValues(firstRows[r], 1, lastRows[r] - firstRows[r] + 1, 1))
  })

  if (Math.min(...firstCols) === Math.max(...lastCols) && Math.min(...firstRows) > 4 && Math.max( ...lastRows) <= sheet.getLastRow()) // If the user has not selected an item, alert them with an error message
  { 
    const numItems = itemValues.length;
    const row = (isNotBlank(sheet.getSheetValues(5, 4, 1, 1)[0][0])) ? 
      Math.max(getLastRowSpecial(sheet.getSheetValues(1, 4, sheet.getMaxRows(), 1)), // SKU column
               getLastRowSpecial(sheet.getSheetValues(1, 8, sheet.getMaxRows(), 1))) // Description column
      + 1: 5;
    sheet.getRange(row, 3, numItems, 6).setNumberFormat('@').setValues(itemValues.map(item => {
      splitDescription = item[0].split(' - ');
      sku = splitDescription.pop();
      uom = splitDescription.pop();
      splitDescription.pop();
      return ['D', sku, 0, '', uom, splitDescription]
    })).offset(0, 2, 1, 1).activate(); // Move to the quantity column
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');

  sheet.getRange(2, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function retrieves the items on the Recently Created and places them on the Item Search sheet.
 * 
 * @author Jarren Ralf
 */
function allItems()
{
  const startTime = new Date().getTime(); // Used for the function runtime
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = SpreadsheetApp.getActiveSheet();
  const recentlyCreatedSheet = spreadsheet.getSheetByName('Recently Created');
  const numItems = recentlyCreatedSheet.getLastRow();
  sheet.getRange(1, 1).clearContent() // Clear the search box
    .offset(4, 0, sheet.getMaxRows() - 4).clearContent() // Clear the previous search
    .offset(0, 0, numItems).setValues(recentlyCreatedSheet.getSheetValues(1, 1, numItems, 1)) // Set the values
    .offset(-3, 7, 1, 1).setValue("Items displayed in order of newest to oldest.") // Tell user items are sorted from newest to oldest
    .offset(0, -1).setValue((new Date().getTime() - startTime)/1000 + " seconds"); // Function runtime
  spreadsheet.toast('PNT\'s most recently created items are being displayed.');
}

/**
 * This function checks if the customer has checked or unchecked their submission button. In both cases, this will lead to the PNT Order Processor spreadsheet to send the appropriate 
 * emails to the customer and PNT employees.
 * 
 * @param {Range} range : The range of the checkbox.
 * @author Jarren Ralf
 */
function checkForOrderSubmission(range)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  range.offset(0, 1).setValue((new Date().getTime() - startTime)/1000 + " seconds").offset(2, -5).setValue(startTime); // Place the timestamp in one of the hidden cells
  SpreadsheetApp.flush(); // Force the change on the spreadsheet first
  SpreadsheetApp.getUi().alert((range.isChecked()) ? 'Your order has been submitted.\n\nThank You!' : 'You have cancelled your order.\n\nYou may make changes and re-submit by selecting the checkbox again.')
}

/**
 * This function
 * 
 * @author Jarren Ralf
 */
function customerSelection(range, spreadsheet)
{
  const selectedCustomer = range.getValue();

  if (isNotBlank(selectedCustomer))
  {
    const customerSheet = spreadsheet.getSheetByName('Customer List');
    range.offset(-1, 4).setValue(customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 2).find(customer => customer[1] === selectedCustomer)[0])
  }
  else
    range.offset(-1, 4).setValue('');
}

/**
 * This function handles the task of deleting items from the users order on the Item Search sheet. 
 * It finds the missing descriptions and it moves the data up to fill in the gap.
 * 
 * @param {Sheet}       sheet      : The Item Search sheet
 * @param {Range}       range      : The active range
 * @param {String[][]}  value      : The values in the range that were editted
 * @param {Number}       row       : The first row that was editted
 * @param {Boolean}  isSingleRow   : Whether or not a single row was editted
 * @param {Boolean} isSingleColumn : Whether or not a single column was editted
 * @author Jarren Ralf
 */
function deleteItemsFromOrder(sheet, range, value, row, isSingleRow, isSingleColumn, spreadsheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  spreadsheet.toast('Checking for possible lines to delete...')
  const numRows = Math.max(getLastRowSpecial(sheet.getSheetValues(1, 4, sheet.getMaxRows(), 1)), getLastRowSpecial(sheet.getSheetValues(1, 8, sheet.getMaxRows(), 1))) - row + 1;

  if (numRows > 0)
  {
    const itemsOrderedRange = sheet.getRange(row, 3, numRows, 6);
    
    if (isSingleRow)
    {
      if (isSingleColumn && !Boolean(value)) // Was a single cell editted?, is the value blank? or is the quantity zero?
      {
        const orderedItems = itemsOrderedRange.getValues();
        orderedItems.shift(); // This is the item that was deleted by the user
        itemsOrderedRange.clearContent()

        if (orderedItems.length > 0)
          itemsOrderedRange.offset(0, 0, orderedItems.length).setValues(orderedItems); // Move the items up to fill in the gap

        spreadsheet.toast('Deleting Complete.')
      }
      else
        spreadsheet.toast('Nothing Deleted.')
    }
    else if (isEveryValueBlank(range.getValues())) // Multiple rows
    {
      const orderedItems = itemsOrderedRange.getValues().filter(description => isNotBlank(description[5])); // Find and remove the blank descriptions
      itemsOrderedRange.clearContent();
      
      if (orderedItems.length > 0)
        itemsOrderedRange.offset(0, 0, orderedItems.length, 6).setValues(orderedItems); // Move the items up to fill in the gaps 

      spreadsheet.toast('Deleting Complete.')
    }
    else
      spreadsheet.toast('Nothing Deleted.')
  }
  else
    spreadsheet.toast('Nothing Deleted.')

  sheet.getRange(2, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * Gets the last row number based on a selected column range values
 *
 * @param {array} range : takes a 2d array of a single column's values
 * @returns {number} : the last row number with a value. 
 */ 
function getLastRowSpecial(range)
{
  for (var row = 0, rowNum = 0, blank = false; row < range.length; row++)
  {
    if (isBlank(range[row][0]) && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if (isNotBlank(range[row][0]))
      blank = false;
  }
  return rowNum;
}

/**
 * This function checks if the given string is blank or not.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is blank or not.
 * @author Jarren Ralf
 */
function isBlank(str)
{
  return str === '';
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
 * This function checks if the given string is not blank or not.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is not blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(str)
{
  return str !== '';
}

/**
 * This function process the imported data.
 * 
 * @param {Event Object} : The event object on an spreadsheet edit.
 * @author Jarren Ralf
 */
function processImportedData(e)
{
  if (e.changeType === 'INSERT_GRID')
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isCustomerData = 4;

    for (var sheet = sheets.length - 1; sheet >= 0; sheet--) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Some sheets in this spreadsheet are OBJECT sheets because they contain full charts
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          sheets[sheet].getRange(1, 1).getValue().toString().includes('Cust #') // A characteristic of the customer data
        ]
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            info[isCustomerData]) 
        {
          spreadsheet.toast('Processing imported data...', '', 60)

          if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of")
          {
            if (info[isCustomerData])
            {
              updateCustomerList(info[numRows], info[numCols], sheets[sheet], spreadsheet)
              spreadsheet.deleteSheet(spreadsheet.getSheetByName('Customer List')) // Delete the old customer list
              sheets[sheet].setName('Customer List').hideSheet().deleteRow(info[numRows]).deleteColumns(3, info[numCols] - 2);
              sheets[sheet].sort(2).setFrozenRows(1);
              sheets[0].getRange(1, 9).activate();
              spreadsheet.toast('Customer List was updated.', 'Import Complete.')
            }
            else // Assume it's a credit note from counter sales
            {
              const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols])
              Logger.log(values)
              spreadsheet.deleteSheet(sheets[sheet]);
              //spreadsheet.toast('Counter Sales Credit Note was successfully imported.', 'Import Complete.')
            }
          }
          
          break;
        }
      }
    }

    // Try and find the file created and delete it
    var file = DriveApp.getFilesByName("Book1.xlsx")

    if (file.hasNext())
      file.next().setTrashed(true)
  }
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the Item List page for the items in question.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function search(e, spreadsheet, sheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  const output = [];
  const searchesOrNot = sheet.getRange(1, 1, 2).clearFormat()                                       // Clear the formatting of the range of the search box
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

    const inventorySheet = spreadsheet.getSheetByName('Item List');
    const data = inventorySheet.getSheetValues(1, 1, inventorySheet.getLastRow(), 1);
    const numSearches = searches.length; // The number searches
    var numSearchWords;

    if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
    {
      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
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
      var numSearchWords_ToNotInclude;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                numSearchWords_ToNotInclude = dontIncludeTheseWords.length - 1;

                for (var l = 0; l <= numSearchWords_ToNotInclude; l++)
                {
                  if (!data[i][0].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                  {
                    if (l === numSearchWords_ToNotInclude)
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
      sheet.getRange('A1').activate() // Move the user back to the seachbox
        .offset(4, 0, sheet.getMaxRows() - 4).clearContent() // Clear content
        .offset(-3, 6, 1, 1).setValue("No results found.\nPlease try again.");
    else
      sheet.getRange('A5').activate() // Move the user to the top of the search items
        .offset(0, 0, sheet.getMaxRows() - 4).clearContent()
        .offset(0, 0, numItems).setValues(output) 
        .offset(-3, 7, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.");

    spreadsheet.toast('Searching Complete.');
  }
  else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
  {
    spreadsheet.toast('Accessing most recently created items...');
    const recentlyCreatedItemsSheet = spreadsheet.getSheetByName('Recently Created');
    const numItems = recentlyCreatedItemsSheet.getLastRow();
    sheet.getRange('A5').activate() // Move the user to the top of the search items
      .offset(0, 0, sheet.getMaxRows() - 4).clearContent()
      .offset(0, 0, numItems).setValues(recentlyCreatedItemsSheet.getSheetValues(1, 1, numItems, 1))
      .offset(-3, 7, 1, 1).setValue("Items displayed in order of newest to oldest.")
    spreadsheet.toast('PNT\'s most recently created items are being displayed.')
  }
  else
  {
    sheet.getRange(5, 1, sheet.getMaxRows() - 4).clearContent() // Clear content 
      .offset(-3, 7, 1, 1).setValue("Invalid search.\nPlease try again.");
    spreadsheet.toast('Invalid Search.');
  }

  sheet.getRange(2, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the Item List page for the items in question.
 * 
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function search_Customer(spreadsheet, sheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  const customers = [];
  const searchesOrNot = sheet.getRange(1, 9).clearFormat()                                          // Clear the formatting of the range of the search box
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
    .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
    .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
    .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
    .getValue().toString().toUpperCase().split(' NOT ')                                             // Split the search string at the word 'not'

  const searches = searchesOrNot[0].split(' OR ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace
  const customerSheet = spreadsheet.getSheetByName('Customer List');

  if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
  {
    const numSearches = searches.length; // The number searches

    // Check if the customer number was entered in the search cell
    if (numSearches === 1 && searches[0].length === 1 && searches[0][0].toString().split(" ").length === 1 && (!Number.isNaN(Number(searches[0][0])) || 
      (searches[0][0].length > 4 && !Number.isNaN(Number(searches[0][0].slice(-4)))) || 
      (searches[0][0].length > 3 && !Number.isNaN(Number(searches[0][0].slice(-3)))) || 
      (searches[0][0].length > 2 && !Number.isNaN(Number(searches[0][0].slice(-2)))) ||
      (searches[0][0].length > 1 && !Number.isNaN(Number(searches[0][0].slice(-1)))) ))
    {
      const customer = customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 2).find(custNum => custNum[0] === searches[0][0]);

      if (customer != undefined) // Customer Number was found
        sheet.getRange(2, 2).setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build()).setValue(customer[1])
          .offset(-1,  4).setValue(customer[0])
          .offset( 1,  3).setValue("1 customer found.")
          .offset(-1, -8).activate();
      else // No customers were found
        sheet.getRange(2, 2).setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build())
          .offset( 0, 7).setValue("No customers found.\nPlease try again.")
          .offset(-1, 0).activate() // Move the user back to the seachbox
        
      spreadsheet.toast('Customer Searching Complete.');
    }
    else // Regular search for the customer name
    {
      const data = customerSheet.getSheetValues(2, 2, customerSheet.getLastRow() - 1, 1);
    
      var numSearchWords;

      if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
      {
        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
          {
            numSearchWords = searches[j].length - 1;

            for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
            {
              if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
              {
                if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                {
                  customers.push(data[i][0]);
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
        var numSearchWords_ToNotInclude;

        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
          {
            numSearchWords = searches[j].length - 1;

            for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
            {
              if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
              {
                if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                {
                  numSearchWords_ToNotInclude = dontIncludeTheseWords.length - 1;

                  for (var l = 0; l <= numSearchWords_ToNotInclude; l++)
                  {
                    if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l]))
                    {
                      if (l === numSearchWords_ToNotInclude)
                      {
                        customers.push(data[i][0]);
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

      const numItems = customers.length;

      if (numItems === 0) // No items were found
        sheet.getRange(2, 2).setValue('').setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build())
          .offset(-1, 4).setValue('')
          .offset( 1, 3).setValue("No customers found.\nPlease try again.")
          .offset(-1, 0).activate() // Move the user back to the seachbox
      else if (numItems !== 1) // More than 1 customer was found
        sheet.getRange(2, 2).setValue('').setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInList(customers).build()).activate() // Move the user to customer data validation
          .offset(-1, 4).setValue('')
          .offset( 1, 3).setValue(numItems + " customers found.");
      else // Only 1 customer was found
        sheet.getRange(2, 2).setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build()).setValue(customers[0])
          .offset(-1,  4).setValue(customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 2).find(customer => customer[1] === customers[0])[0])
          .offset( 1,  3).setValue("1 customer found.")
          .offset(-1, -8).activate();
        
      spreadsheet.toast('Customer Searching Complete.');
    }
  }
  else
  {
    sheet.getRange(2, 2).setValue('').setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build()).activate() // Move the user to customer data validation
      .offset(-1, 4).setValue('')
      .offset( 1, 3).setValue("All customers displayed alphabetically.");
    spreadsheet.toast('All customers displayed in Data Validation alphabetically.');
  }

  sheet.getRange(1, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This fucntion...
 */
function updateCustomerList()
{

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
/**
 * This function handles the on edit events in this spreadsheet pertaining to the Item Search sheet only (all other sheets will be protected).
 * This function is looking for the user searching for items and it is making appropriate changes to the data when a user deletes items from their order.
 * 
 * @param {Event Object} e : The event object
 */
function installedOnEdit(e)
{ 
  const range = e.range;
  const col = range.columnStart;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const isSingleRow = row == rowEnd;
  const isSingleColumn = col == range.columnEnd;
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();

  if (isSingleColumn && sheet.getSheetName() === 'Item Search')
  {
    if (row == 1)
    {
      if (col == 1 && (rowEnd == null || rowEnd == 2 || isSingleRow)) // Item Search
        search(e, spreadsheet, sheet, false);
      else if (col == 10 && (rowEnd == null || isSingleRow)) // Customer Search
        search_Customer(spreadsheet, sheet);
      else if (col == 2) // Price Selection
        priceSelection(range, sheet, spreadsheet);
    }
    else if (row == 2 && col == 2)
      customerSelection(range, spreadsheet);
    else if (row > 4)
    {
      if (col == 1 && rowEnd >= row)
        search(e, spreadsheet, sheet, true);
      else if (col == 9)
        deleteItemsFromOrder(sheet, range, range.getValue(), row, isSingleRow, isSingleColumn, spreadsheet);
    }
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
  SpreadsheetApp.getUi().createMenu('Export')
    .addItem('Add PNT Delivery Charge ($25)',  'addFreight_PntDelivery')
    .addItem('Add LMFF Shipping Charge', 'addFreight_LMFF')
    .addSeparator()
    .addItem('Complete', 'completeOrder')
    .addItem('Complete (Fully shipped | No BOs)', 'completeOrder_FullyShipped')
    .addSeparator()
    .addItem('Clear Export Page', 'clearExport').addToUi();
}

/**
 * This function adds a freight charge to the order.
 * 
 * @param {Boolean} isPntDelivery : Whether the ordwr is being shipped via PNT Delivery or not.
 * @author Jarren Ralf
 */
function addFreight(isPntDelivery)
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const maxRows = sheet.getMaxRows();
  const today = Utilities.formatDate(new Date(), SpreadsheetApp.getActive(), 'dd MMM')
  const lastRow = Math.max(getLastRowSpecial(sheet.getSheetValues(1,  3, maxRows, 1)),
                           getLastRowSpecial(sheet.getSheetValues(1,  4, maxRows, 1)), 
                           getLastRowSpecial(sheet.getSheetValues(1,  9, maxRows, 1)), 
                           getLastRowSpecial(sheet.getSheetValues(1, 10, maxRows, 1))) - 3;

  sheet.getRange(4, 6).setValue((isPntDelivery) ? 'PNT DELIVERY' : 'LOWER MAINLAND F.F.')
    .offset(lastRow, -3, 1, 7).setValues((isPntDelivery) ? 
      [['D', 'FREIGHT', '25.00', 1, 1, 'EACH', 'PNT Delivery on ' + today + ' to office']] : 
      [['D', 'FREIGHT', '25.00', 1, 1, 'EACH', 'Shipped via LMFF on ' + today + ' to store']])
    .offset(0, 2, 1, 1).activate();
}

/**
 * This function adds a PNT Delivery freight charge to the order at a value of $25.00.
 * 
 * @author Jarren Ralf
 */
function addFreight_LMFF()
{
  addFreight(false);
}

/**
 * This function adds a PNT Delivery freight charge to the order at a value of $25.00.
 * 
 * @author Jarren Ralf
 */
function addFreight_PntDelivery()
{
  addFreight(true)
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
    const maxRow = sheet.getMaxRows();
    const row = (isNotBlank(sheet.getSheetValues(5, 4, 1, 1)[0][0])) ? 
      Math.max(getLastRowSpecial(sheet.getSheetValues(1, 4, maxRow, 1)), // SKU column
               getLastRowSpecial(sheet.getSheetValues(1, 8, maxRow, 1))) // Description column
      + 1: 5;
    sheet.getRange(row, 3, numItems, 7).setNumberFormat('@').setValues(itemValues.map(item => {
      splitDescription = item[0].split(' - ');
      sku = splitDescription.pop();
      uom = splitDescription.pop();
      splitDescription.pop();
      splitDescription.pop();
      return ['D', sku, 0, '', 0, uom, splitDescription.join(' - ')]
    })).offset(0, 3, 1, 1).activate() // Move to the quantity column
    .offset(1 - row, 2, 1, 1).setValue((new Date().getTime() - startTime)/1000 + " seconds");
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');
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
    .offset(4, 0, sheet.getMaxRows() - 4).clearContent().setWrap(true) // Clear the previous search
    .offset(0, 0, numItems).setValues(recentlyCreatedSheet.getSheetValues(1, 1, numItems, 1)) // Set the values
    .offset(-3, 9, 1, 1).setValue("Items displayed in order of newest to oldest.") // Tell user items are sorted from newest to oldest
    .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds"); // Function runtime
  spreadsheet.toast('PNT\'s most recently created items are being displayed.');
}

/**
 * This function clears the export sheet and then sends Adrian a courtesy email letting him know that the import template for Adagio OrderEntry has changed.
 * 
 * @author Jarren Ralf
 */
function clearExport()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const exportSheet = spreadsheet.getActiveSheet();

  try
  {
    if (exportSheet.getSheetName() !== 'Export')
    {
      spreadsheet.getSheetByName('Export').activate();
      Browser.msgBox('You must be on the Export sheet in order to clear it.')
    }
    else
      exportSheet.clear(); 
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error);
  }
}

/**
 * This function places the current Order on the Export page for importing.
 * 
 * @param {Boolean} isFullyShipped : Whether the order is fully shipped, i.e. order quantity = shipped quantity
 * @author Jarren Ralf
 */
function completeOrder(isFullyShipped)
{
  try
  {
    const searchSheet = SpreadsheetApp.getActiveSheet();

    if (searchSheet.getSheetName() !== 'Item Search')
    {
      spreadsheet.getSheetByName('Item Search').activate();
      Browser.msgBox('Please return to the Dashboard to run this function.')
    }
    else
    {
      var exportData_WithDiscountedPrices = [];
      const maxRow = searchSheet.getMaxRows() - 3;
      const numRows = Math.max(getLastRowSpecial(searchSheet.getSheetValues(4,  4, maxRow, 1)), 
                               getLastRowSpecial(searchSheet.getSheetValues(4,  9, maxRow, 1)), 
                               getLastRowSpecial(searchSheet.getSheetValues(4, 10, maxRow, 1)));
      const numCols = 8;
      const orderRange = searchSheet.getRange(4, 3, numRows, numCols);
      const shippedQtyIndex = (isFullyShipped) ? 3 : 4;

      orderRange.getValues().map(item => {
        if (item[0] === 'H')
          exportData_WithDiscountedPrices.push(['H', item[1], item[2], item[3], item[4], item[5]])
        else if (item[0] === 'I')
          exportData_WithDiscountedPrices.push(['I', item[1], '', '', '', ''])
        else if (item[0] === 'D')
        {
          item[1] = item[1].toString().trim().toUpperCase(); // Make the SKU uppercase

          if (isNotBlank(item[1])) // SKU is not blank
            if (isNotBlank(item[3])) // Order quantity is not blank
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
                exportData_WithDiscountedPrices.push(['D', item[1], item[2], item[3], item[shippedQtyIndex], item[6].toString().substring(0, 40)])
              else // Order quantity is not a valid number
                exportData_WithDiscountedPrices.push(
                  ['D', item[1], item[2], 0, 0, item[6].toString().substring(0, 40)], 
                  ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', '', '', '']
                )
            else // The order quantity is blank (while SKU is not)
              exportData_WithDiscountedPrices.push(
                ['D', item[1], item[2], 0, 0, item[6].toString().substring(0, 40)],
                ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '', '', '']
              )
          else // The SKU is blank
            if (isNotBlank(item[3])) // Order quantity is not blank
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, item[3], item[shippedQtyIndex], item[6].toString().substring(0, 40)], 
                  ...('Description: ' + item[6] + ' - ' + item[5]).toString().match(/.{1,75}/g).map(c => ['C', c, '', '', '', ''])
                )
              else // Order quantity is not a valid number
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0, 0, item[6].toString().substring(0, 40)], 
                  ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', '', '', ''], 
                  ...('Description: ' + item[6] + ' - ' + item[5]).toString().match(/.{1,75}/g).map(c => ['C', c, '', '', '', ''])
                )
            else if (isNotBlank(item[6]))// Description is not blank (but SKU and quantity are)
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0, 0, item[6].toString().substring(0, 40)], 
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '', '', ''],
                  ...('Description: ' + item[6] + ' - ' + item[5]).toString().match(/.{1,75}/g).map(c => ['C', c, '', '', '', ''])
                )

          if (isNotBlank(item[7])) // There are notes for the current line
            exportData_WithDiscountedPrices.push(...('Notes: ' + item[7]).match(/.{1,75}/g).map(c => ['C', c, '', '', '', '']))
        }
        else // There was no line indicator
        {
          item[1] = item[1].toString().trim().toUpperCase(); // Make the SKU uppercase

          if (isNotBlank(item[1])) // SKU is not blank
          {
            if (isNotBlank(item[3])) // Order quantity is not blank
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
                exportData_WithDiscountedPrices.push(['D', item[1], item[2], item[3], item[shippedQtyIndex], item[6].toString().substring(0, 40)])
              else // Order quantity is not a valid number
                exportData_WithDiscountedPrices.push(
                  ['D', item[1], item[2], 0, 0, item[6].toString().substring(0, 40)], 
                  ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', '', '', '']
                )
            else // The order quantity is blank (while SKU is not)
              exportData_WithDiscountedPrices.push(
                ['D', item[1], item[2], 0, 0, item[6].toString().substring(0, 40)],
                ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '', '', '']
              )
          }
          else // The SKU is blank
          {
            if (isNotBlank(item[3])) // Order quantity is not blank
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, item[3], item[shippedQtyIndex], item[6].toString().substring(0, 40)], 
                  ...('Description: ' + item[6] + ' - ' + item[5]).toString().match(/.{1,75}/g).map(c => ['C', c, '', '', '', ''])
                )
              else // Order quantity is not a valid number
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0, 0, item[6].toString().substring(0, 40)], 
                  ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', '', '', ''], 
                  ...('Description: ' + item[6] + ' - ' + item[5]).toString().match(/.{1,75}/g).map(c => ['C', c, '', '', '', ''])
                )
            else if (isNotBlank(item[6])) // Description is not blank (but SKU and quantity are)
              exportData_WithDiscountedPrices.push(
                ['D', 'MISCITEM', 0, 0, 0, item[6].toString().substring(0, 40)], 
                ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '', '', ''],
                ...('Description: ' + item[6] + ' - ' + item[5]).toString().match(/.{1,75}/g).map(c => ['C', c, '', '', '', ''])
              )
          }

          if (isNotBlank(item[7])) // There are notes for the current line
            exportData_WithDiscountedPrices.push(...('Notes: ' + item[7]).match(/.{1,75}/g).map(c => ['C', c, '', '', '', '']))
        }
      })

      orderRange.offset(1, 0, numRows - 1, numCols).clearContent() // Customer Order
        .offset(-1,  3, 1, 3)
          .setValues([['PNT DELIVERY', '', '']])  // PNT Delivery, Invoice Discount Percentage, and Comment 1
        .offset(-3, -4, 1, 1).setValue('')       // Pricing Selection
        .offset( 1,  0).setValue('')             // Customer Name
        .offset( 0,  7).setValue('')             // PO Number
        .offset(-1, -3).setValue('')             // Customer #
        .offset( 1,  4).setValue('')             // Message Display
        .offset(-1,  0).setValue('').activate(); // Customer Search
      const exportSheet = SpreadsheetApp.getActive().getSheetByName('Export');
      const lastRow = exportSheet.getLastRow() + 1;
      const ranges = [[],[],[]];
      const backgroundColours = [
        '#c9daf8', // Make the header rows blue
        '#fcefe1', // Make the comment rows orange
        '#e0d5fd'  // Make the instruction comment rows purple
      ];

      exportData_WithDiscountedPrices.map((h, r) => 
        h = (h[0] !== 'H') ? (h[0] !== 'C') ? (h[0] !== 'I') ? false : 
        ranges[2].push('A' + (r + lastRow) + ':F' + (r + lastRow)) : // Instruction comment rows purple
        ranges[1].push('A' + (r + lastRow) + ':F' + (r + lastRow)) : // Comment rows orange
        ranges[0].push('A' + (r + lastRow) + ':F' + (r + lastRow))   // Header rows blue
      )

      ranges.map((rngs, r) => (rngs.length !== 0) ? exportSheet.getRangeList(rngs).setBackground(backgroundColours[r]) : false); // Set the appropriate background colours
      exportSheet.getRange(lastRow, 1, exportData_WithDiscountedPrices.length, 6).setNumberFormat('@').setValues(exportData_WithDiscountedPrices)
        .offset(-1*lastRow + 1, 0, exportSheet.getLastRow(), 6).activate();
    }
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * This function places the current Order on the Export page for importing, and makes the shipped quantities match the order quantity.
 * 
 * @author Jarren Ralf
 */
function completeOrder_FullyShipped()
{
  completeOrder(true)
}

/**
 * This function takes the selected customer name from the drop down and retreives the corresponding Customer number.
 * 
 * @param {Range} range : The range of the data validation where the vendor name is selected.
 * @param {Spreadsheet} : The active spreadsheet.
 * @author Jarren Ralf
 */
function customerSelection(range, spreadsheet)
{
  const selectedCustomer = range.getValue();

  if (isNotBlank(selectedCustomer))
  {
    const customerSheet = spreadsheet.getSheetByName('Customer List');
    range.offset(-1, 4).setValue(customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 2).find(customer => customer[1] === selectedCustomer)[0]).offset(0, -5).activate()
  }
  else
    range.offset(-1, 4).setValue('');
}

/**
 * This function creates the trigger for updating the items daily and for the installed onEdit and Change triggers.
 * 
 * @author Jarren Ralf
 */
function triggers_CreateAll()
{
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('updateItems').timeBased().everyDays(1).atHour(23).create();
  ScriptApp.newTrigger('updateUPCs').timeBased().everyDays(1).atHour(23).create();
  ScriptApp.newTrigger('onChange').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('installedOnEdit').forSpreadsheet(ss).onEdit().create();
}

/**
 * This function creates the trigger for updating the items daily and for the installed onEdit and Change triggers.
 * 
 * @author Jarren Ralf
 */
function triggers_DeleteAll()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger));
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
  const maxRow = sheet.getMaxRows();
  const numRows = Math.max(getLastRowSpecial(sheet.getSheetValues(1, 4, maxRow, 1)), getLastRowSpecial(sheet.getSheetValues(1, 9, maxRow, 1))) - row + 1;

  if (numRows > 0)
  {
    const itemsOrderedRange = sheet.getRange(row, 3, numRows, 7);
    
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
      const orderedItems = itemsOrderedRange.getValues().filter(description => isNotBlank(description[6])); // Find and remove the blank descriptions
      itemsOrderedRange.clearContent();
      
      if (orderedItems.length > 0)
        itemsOrderedRange.offset(0, 0, orderedItems.length, 7).setValues(orderedItems); // Move the items up to fill in the gaps 

      spreadsheet.toast('Deleting Complete.')
    }
    else
      spreadsheet.toast('Nothing Deleted.')
  }
  else
    spreadsheet.toast('Nothing Deleted.')

  sheet.getRange(1, 8).setValue((new Date().getTime() - startTime)/1000 + " seconds");
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
 * This function returns true if the presented number is a UPC-A, false otherwise.
 * 
 * @param {Number} upcNumber : The UPC-A number
 * @returns Whether the given value is a UPC-A or not
 * @author Jarren Ralf
 */
function isUPC_A(upcNumber)
{
  for (var i = 0, sum = 0, upc = upcNumber.toString(); i < upc.length - 1; i++)
    sum += (i % 2 === 0) ? Number(upc[i])*3 : Number(upc[i]);

  return upc.endsWith(Math.ceil(sum/10)*10 - sum) && upc.length === 12;
}

/**
 * This function returns true if the presented number is a EAN_13, false otherwise.
 * 
 * @param {Number} upcNumber : The EAN_13 number
 * @returns Whether the given value is a EAN_13 or not
 * @author Jarren Ralf
 */
function isEAN_13(upcNumber)
{
  for (var i = 0, sum = 0, upc = upcNumber.toString(); i < upc.length - 1; i++)
    sum += (i % 2 === 0) ? Number(upc[i]) : Number(upc[i])*3;

  return upc.endsWith(Math.ceil(sum/10)*10 - sum) && upc.length === 13;
}

/**
 * This function changes the pricing for the whole order based on the user's selection of pricing structure.
 * 
 * @param   {Range}       range     : The range of the price selection.
 * @param   {Sheet}       sheet     : The active sheet.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function priceSelection(range, sheet, spreadsheet)
{
  const selectedPricing = range.getValue();

  if (isNotBlank(selectedPricing))
  {
    spreadsheet.toast('Computing price change...')

    switch (selectedPricing)
    {
      case 'Retail Pricing':
        var price = 1;
        break;
      case 'Guide Pricing':
        var price = 2;
        break;
      case 'Lodge Pricing':
        var price = 3;
        break;
      case 'Wholesale Pricing':
        var price = 4;
        break;
    }

    const discountSheet = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs').getSheetByName('Discount Percentages')
    const discounts = discountSheet.getSheetValues(2, 11, discountSheet.getLastRow() - 1, 5);
    const BASE_PRICE = 1;
    var itemPricing;

    const maxRows = sheet.getMaxRows();
    const numRows = getLastRowSpecial(sheet.getSheetValues(5, 4, maxRows, 1));

    if (numRows > 0)
    {
      const itemRange = sheet.getRange(5, 4, getLastRowSpecial(sheet.getSheetValues(5, 4, maxRows, 1)), 2);

      const order = itemRange.getValues().map(item => {
        if (item[0] !== 'FREIGHT')
        {
          itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[0]); // Find the item pricing on the discount sheet

          if (itemPricing != undefined && itemPricing[BASE_PRICE] != 0) // SKU is assumed to be valid
            item[1] = (price !== 1) ? (itemPricing[BASE_PRICE]*(100 - itemPricing[price])/100).toFixed(2) : itemPricing[price];
        }
        else
          item[1] = '25.00';

        return item
        
      }) 

      itemRange.setValues(order);
      spreadsheet.toast('Price change complete.')
    }
    else
    {
      range.offset(0, -1).activate();
      spreadsheet.toast('Please add some items to the order.')
    }
  }
}

/**
 * This function checks processes the recently created tab and converts the information into a new order that can be imported into Adagio. It is assumed to be a countersales credit note in this case.
 * 
 * @param {Number} numRows : The number of rows from the imported excel file.
 * @param {Number} numCols : The number of columns from the imported excel file.
 * @param {Sheet} newSheet : The new sheet that that was just created on import that contains the data from the excel file.
 * @param {Sheet} itemSearchSheet : The item search search which is used to create new orders.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf  
 */
function processImportedCreditNote(numRows, numCols, newSheet, itemSearchSheet, spreadsheet)
{
  const values = newSheet.getSheetValues(1, 1, numRows, numCols);
  const custNum = values[2][24];
  const poNum = values[9][23];
  const invoiceDiscountPercentage = (values[values.length - 8][23].substring(2) != (Number(values[values.length - 10][23]) + Number(values[values.length - 11][23]) + Number(values[values.length - 13][23]))) ? 
    Math.round(100 - (Number(values[values.length - 8][23].substring(2)) + Number(values[values.length - 10][23]) + Number(values[values.length - 11][23]))/Number(values[values.length - 13][23])*-100) : '';
  const customerSheet = spreadsheet.getSheetByName('Customer List');
  var custName = customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 2).find(custNumber => custNumber[0] === custNum);
  custName = (custName != undefined) ? custName[1] : '';
  const items = values.filter((val, idx) => isNotBlank(val[16]) && (idx - 11)%40 > 0 && (idx - 11)%40 < 24); // Use the unit of measure column to remove unnecessary rows
  items.pop();
  items.pop();
  items.pop();
  const exportValues = items.map(item => ['D', item[3], item[18], (-1*Number(item[0])).toString(), (-1*Number(item[0])).toString(), item[16], item[5]]);

  itemSearchSheet.getRange(1, 2).setValue('')
    .offset(0,  4).setValue(custNum)
    .offset(1, -4).setDataValidation(itemSearchSheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build()).setValue('\'' + custName)   
    .offset(0,  7).setValue(poNum)
    .offset(2, -2).setValue(invoiceDiscountPercentage)
    .offset(1, -4, itemSearchSheet.getMaxRows() - 4, 8).clearContent()
    .offset(0, 0, exportValues.length, 7).setNumberFormat('@').setValues(exportValues).activate();
  spreadsheet.deleteSheet(newSheet);
  spreadsheet.toast('Counter Sales Credit Note was successfully imported.', 'Import Complete.')
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
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isCustomerData = 4, isCounterSalesInvoice = 5, isCounterSalesCreditNote = 6, fileName;

    for (var sheet = sheets.length - 1; sheet >= 0; sheet--) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Some sheets in this spreadsheet are OBJECT sheets because they contain full charts
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          (sheets[sheet].getLastColumn() != 0) ? sheets[sheet].getSheetValues(1,  1, 1, sheets[sheet].getLastColumn())[0].includes('Telephone') : false, // A characteristic of the customer data
          (sheets[sheet].getMaxColumns() > 23) ? sheets[sheet].getSheetValues(1,  1, 1, sheets[sheet].getLastColumn())[0].includes('Invoice') : false, // A way of distinguishing countersales invoices
          (sheets[sheet].getMaxColumns() > 23) ? sheets[sheet].getSheetValues(1,  1, 1, sheets[sheet].getLastColumn())[0].includes('Credit') : false // A way of distinguishing countersales credit notes 
        ]

        fileName = sheets[sheet].getSheetName()
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            (info[maxRow] === 1000 && info[maxCol] === info[numCols] && info[numRows] !== 0) ||
            info[isCustomerData]) 
        {
          spreadsheet.toast('Processing imported data...', '', 60)

          if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of")
          {
            if (info[isCustomerData])
              updateCustomerList(info[numRows], info[numCols], sheets[sheet], sheets, spreadsheet);
            else if (info[isCounterSalesCreditNote]) // Assume it's a credit note from counter sales
              processImportedCreditNote(info[numRows], info[numCols], sheets[sheet], sheets[0], spreadsheet)
            else if (info[isCounterSalesInvoice])
              processImportedInvoice(info[numRows], info[numCols], sheets[sheet], sheets[0], spreadsheet)
          }
          
          break;
        }
      }
    }

    // Try and find the file created and delete it
    var file1 = DriveApp.getFilesByName("Book1.xlsx")
    var file2 = DriveApp.getFilesByName(fileName + ".xlsx")

    if (file1.hasNext())
      file1.next().setTrashed(true)

    if (file2.hasNext())
      file2.next().setTrashed(true)
  }
}

/**
 * This function checks processes the recently created tab and converts the information into a new order that can be imported into Adagio. It is assumed to be a countersale invocie in this case.
 * 
 * @param {Number} numRows : The number of rows from the imported excel file.
 * @param {Number} numCols : The number of columns from the imported excel file.
 * @param {Sheet} newSheet : The new sheet that that was just created on import that contains the data from the excel file.
 * @param {Sheet} itemSearchSheet : The item search search which is used to create new orders.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf  
 */
function processImportedInvoice(numRows, numCols, newSheet, itemSearchSheet, spreadsheet)
{
  const values = newSheet.getSheetValues(1, 1, numRows, numCols)
  const custNum = values[1][23];
  const poNum = values[8][22].split('PO #: ').pop();
  const invoiceDiscountPercentage = (values[values.length - 11][19] !== 'Less') ? '' : 
                                    (values[values.length - 11][23] != 0) ? Math.round(Number(values[values.length - 11][23])/Number(values[values.length - 12][23])*100).toString() : '';
  const customerSheet = spreadsheet.getSheetByName('Customer List');
  var custName = customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 2).find(custNumber => custNumber[0] === custNum)
  custName = (custName != undefined) ? custName[1] : '';
  const items = values.filter((val, idx) => isNotBlank(val[15]) && (idx - 10)%38 > 0 && (idx - 10)%38 < 24) // Use the unit of measure column to remove unnecessary rows
  const exportValues = items.map(item => ['D', item[3], item[20], item[0].toString(), item[0].toString(), item[15], item[5]]);

  itemSearchSheet.getRange(1, 2).setValue('')
    .offset(0,  4).setValue(custNum)
    .offset(1, -4).setDataValidation(itemSearchSheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build()).setValue('\'' + custName)
    .offset(0,  7).setValue(poNum)
    .offset(2, -2).setValue(invoiceDiscountPercentage)
    .offset(1, -4, itemSearchSheet.getMaxRows() - 4, 8).clearContent()
    .offset(0, 0, exportValues.length, 7).setNumberFormat('@').setValues(exportValues).activate();
  spreadsheet.deleteSheet(newSheet);
  spreadsheet.toast('Counter Sales Invoice was successfully imported.', 'Import Complete.')
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the Item List page for the items in question.
 * 
 * @param {Event Object}       e           : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}   spreadsheet      : The spreadsheet that is being edited
 * @param    {Sheet}         sheet         : The sheet that is being edited
 * @param   {Boolean} isMultipleItemSearch : Whether the user pasted multiple skus in the description column to search multiple items simultaneously
 * @author Jarren Ralf 
 */
function search(e, spreadsheet, sheet, isMultipleItemSearch)
{
  const startTime = new Date().getTime(); // Used for the function runtime

  if (isMultipleItemSearch) // Check and make sure only a single row is being edited
  {
    spreadsheet.toast('Searching...')                                                           
    const values = e.range.getValues().filter(blank => isNotBlank(blank[0]))
    sheet.getRange(1, 1, 2, 1).clearContent(); // Clear the search bar

    if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
    {
      const inventorySheet = spreadsheet.getSheetByName('Item List');
      const data = inventorySheet.getSheetValues(1, 1, inventorySheet.getLastRow(), 1);
      var someSKUsNotFound = false, skus;

      if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
      {
        skus = values.map(item => {
        
          for (var i = 0; i < data.length; i++)
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item[0].toString().split(' - ').pop().toUpperCase())
              return data[i];
  
          someSKUsNotFound = true;

          return ['SKU Not Found: ' + item[0].toString().split(' - ').pop().toUpperCase()]
        });
      }
      else if (values[0][0].toString().includes('-')) // The SKU contains dashes because that's the convention from Adagio
      {
        skus = values.map(sku => sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).map(item => {
        
          for (var i = 0; i < data.length; i++)
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item.toString().toUpperCase())
              return data[i];

          someSKUsNotFound = true;

          return ['SKU Not Found: ' + item]
        });
      }
      else // The regular plain SKUs are being pasted
      {
        skus = values.map(item => {
        
          for (var i = 0; i < data.length; i++)
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item[0].toString().toUpperCase())
              return data[i];

          someSKUsNotFound = true;

          return ['SKU Not Found: ' + item[0]]
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
        const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(['#ffe599']), new Array(numSkusFound).fill(['white'])]); // Concatenate all of the item values as a 2-D array

        if (numItems === 0) // No items were found
          sheet.getRange('A1').activate() // Move the user back to the seachbox
            .offset( 4,  0, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setWrap(true) // Clear content
            .offset(-3,  9, 1, 1).setValue("No results found.\nPlease try again.")
            .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
        else
          sheet.getRange('A5') // Move the user to the top of the search items
            .offset( 0, 0, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setFontColor('#434343').setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left').setWrap(true)
              .setBorder(false, false, false, true, false, false, '#1155cc',SpreadsheetApp.BorderStyle.SOLID_THICK)
            .offset( 0,  0, numItems).setValues(items).setBackgrounds(colours).setFontFamily('Arial').setFontWeight('bold')
            .offset(-3,  9, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.")
            .office(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
            .offset((numSkusFound != 0) ? numSkusNotFound + 4 : 4, -7, (numSkusFound != 0) ? numSkusFound : numSkusNotFound, 1).activate();
      }
      else // All SKUs were succefully found
      {
        const numItems = skus.length

        if (numItems === 0) // No items were found
          sheet.getRange('A1').activate() // Move the user back to the seachbox
            .offset( 4, 0, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setWrap(true) // Clear content
            .offset(-3, 9, 1, 1).setValue("No results found.\nPlease try again.")
            .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
        else
          sheet.getRange('A5') // Move the user to the top of the search items
            .offset( 0, 0, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setFontColor('#434343').setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left').setWrap(true)
              .setBorder(false, false, false, true, false, false, '#1155cc',SpreadsheetApp.BorderStyle.SOLID_THICK)
            .offset( 0, 0, numItems).setValues(skus).activate() 
            .offset(-3, 9, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.")
            .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
      }
    }
    spreadsheet.toast('Searching Complete.');
  }
  else
  {
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

      const numSearches = searches.length; // The number searches
      var isBarcodeScanned = false;
      var numSearchWords;

      if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
      {
        if (/^\d+$/.test(searches[0][0]) && (isUPC_A(searches[0][0]) || isEAN_13(searches[0][0])) && numSearches === 1 && searches[0].length == 1) // Check if a barcode was scanned in the cell
        {
          const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database')
          const upcs = upcDatabaseSheet.getSheetValues(1, 1, upcDatabaseSheet.getLastRow(), 1)
          var l = 0; // Lower-bound
          var u = upcs.length - 1; // Upper-bound
          var m = Math.ceil((u + l)/2) // Midpoint
          searches[0][0] = parseInt(searches[0][0])
          isBarcodeScanned = true;

          while (l < m && u > m) // Loop through the UPC codes using the binary search algorithm
          {
            if (searches[0][0] < parseInt(upcs[m][0]))
              u = m;   
            else if (searches[0][0] > parseInt(upcs[m][0]))
              l = m;
            else // UPC code was found!
            {
              const splitDescription = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0].toString().toUpperCase().split(' - ')
              const sku = splitDescription.pop();
              const uom = splitDescription.pop();
              splitDescription.pop();
              splitDescription.pop();
              output.push(['D', sku, 0, '', 0, uom, splitDescription.join(' - ')]);

              var newItemRow = (isNotBlank(sheet.getSheetValues(5, 4, 1, 1)[0][0])) ? 
                  Math.max(getLastRowSpecial(sheet.getSheetValues(1, 4, sheet.getMaxRows(), 1)), // SKU column
                  getLastRowSpecial(sheet.getSheetValues(1, 9, sheet.getMaxRows(), 1))) // Description column
                + 1 : 5;
                
              break; // Item was found, therefore stop searching
            }
              
            m = Math.ceil((u + l)/2) // Midpoint
          }
        }
        else
        {
          const inventorySheet = spreadsheet.getSheetByName('Item List');
          const data = inventorySheet.getSheetValues(1, 1, inventorySheet.getLastRow(), 1);

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
      }
      else // The word 'not' was found in the search string
      {
        const inventorySheet = spreadsheet.getSheetByName('Item List');
        const data = inventorySheet.getSheetValues(1, 1, inventorySheet.getLastRow(), 1);
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
          .offset( 4, 0, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setWrap(true) // Clear content
          .offset(-3, 9, 1, 1).setValue("No results found.\nPlease try again.")
          .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
      else if (isBarcodeScanned)
        sheet.getRange(5, 1, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setWrap(true)
          .offset(-3, 9, 1, 1).setValue("Barcode found.")
          .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
          .offset(newItemRow - 1, -5, 1, 7).setValues(output) 
          .offset(0, 3, 1, 1).activate()
      else
        sheet.getRange('A5').activate() // Move the user to the top of the search items
          .offset( 0, 0, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setWrap(true)
          .offset( 0, 0, numItems).setValues(output) 
          .offset(-3, 9, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.")
          .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")

      spreadsheet.toast('Searching Complete.');
    }
    else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
    {
      spreadsheet.toast('Accessing most recently created items...');
      const recentlyCreatedItemsSheet = spreadsheet.getSheetByName('Recently Created');
      const numItems = recentlyCreatedItemsSheet.getLastRow();
      sheet.getRange('A5').activate() // Move the user to the top of the search items
        .offset( 0, 0, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setWrap(true)
        .offset( 0, 0, numItems).setValues(recentlyCreatedItemsSheet.getSheetValues(1, 1, numItems, 1))
        .offset(-3, 9, 1, 1).setValue("Items displayed in order of newest to oldest.")
        .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
      spreadsheet.toast('PNT\'s most recently created items are being displayed.')
    }
    else
    {
      sheet.getRange(5, 1, sheet.getMaxRows() - 4).clearContent().setBackground('#cccccc').setWrap(true) // Clear content 
        .offset(-3, 9, 1, 1).setValue("Invalid search.\nPlease try again.")
        .offset(-1, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds")
      spreadsheet.toast('Invalid Search.');
    }
  }
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
  const searchesOrNot = sheet.getRange(1, 10).clearFormat()                                         // Clear the formatting of the range of the search box
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
          .offset( 1,  4).setValue("1 customer found.")
          .offset(-1, -9).activate()
          .offset( 0,  7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
      else // No customers were found
        sheet.getRange(2, 2).setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build())
          .offset( 0,  8).setValue("No customers found.\nPlease try again.")
          .offset(-1,  0).activate() // Move the user back to the seachbox
          .offset( 0, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds");
        
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
          .offset(-1,  4).setValue('')
          .offset( 1,  4).setValue("No customers found.\nPlease try again.")
          .offset(-1,  0).activate() // Move the user back to the seachbox
          .offset( 0, -2).setValue((new Date().getTime() - startTime)/1000 + " seconds");
      else if (numItems !== 1) // More than 1 customer was found
        sheet.getRange(2, 2).setValue('').setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInList(customers).build()).activate() // Move the user to customer data validation
          .offset(-1,  4).setValue('')
          .offset( 1,  4).setValue(numItems + " customers found.")
          .offset( 0, -8).activate()
          .offset(-1,  6).setValue((new Date().getTime() - startTime)/1000 + " seconds");
      else // Only 1 customer was found
        sheet.getRange(2, 2).setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build()).setValue(customers[0])
          .offset(-1,  4).setValue(customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 2).find(customer => customer[1] === customers[0])[0])
          .offset( 1,  4).setValue("1 customer found.")
          .offset(-1, -9).activate()
          .offset( 0,  7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
        
      spreadsheet.toast('Customer Searching Complete.');
    }
  }
  else
  {
    sheet.getRange(2, 2).setValue('').setDataValidation(sheet.getRange(2, 2).getDataValidation().copy().requireValueInRange(customerSheet.getRange('$B$2:$B')).build()).activate() // Move the user to customer data validation
      .offset(-1, 4).setValue('')
      .offset( 1, 4).setValue("All customers displayed alphabetically.");
    spreadsheet.toast('All customers displayed in Data Validation alphabetically.');
  }
}

/**
* Sorts data by the created date of the product for the richmond spreadsheet.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCreatedDate(a, b)
{
  return (a[1] === b[1]) ? 0 : (a[1] < b[1]) ? 1 : -1;
}

/**
 * This function sorts the UPC Codes in numerical order.
 * 
 * @author Jarren Ralf
 */
function sortUPCsNumerically(a, b)
{
  return parseInt(a[0]) - parseInt(b[0]);
}

/**
 * This function manages the imported list of Customer names and numbers and puts that information on the hidden Customer List sheet.
 * 
 * @param   {Number}     numRows    : The number of rows on the imported Customer sheet
 * @param   {Number}     numCols    : The number of columns on the imported Customer sheet
 * @param   {Sheet}       sheet     : The imported sheet (The new Customer list)
 * @param   {Sheet[]}     sheets    : All of the sheets of the spreadsheet
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 */
function updateCustomerList(numRows, numCols, sheet, sheets, spreadsheet)
{
  spreadsheet.deleteSheet(spreadsheet.getSheetByName('Customer List')) // Delete the old customer list
  sheet.setName('Customer List').hideSheet().deleteRow(numRows).deleteColumns(3, numCols - 2);
  sheet.sort(2).setFrozenRows(1);
  sheets[0].getRange(1, 10).activate();
  spreadsheet.toast('Customer List was updated.', 'Import Complete.')
}

/**
 * This function updates all of the items daily.
 * 
 * @author Jarren Ralf
 */
function updateItems()
{
  var d,  itemList = [];
  const spreadsheet = SpreadsheetApp.getActive();
  const sortedItems = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString()).map(item => {
    itemList.push([item[1]]);
    d = item[6].split('.');                           // Split the date at the "."
    item[6] = new Date(d[2],d[1] - 1,d[0]).getTime(); // Convert the date sting to a striong object for sorting purposes
  
    return [item[1], item[6]];
  }).sort(sortByCreatedDate).sort(sortByCreatedDate).map(descrip => [descrip[0]])

  // Remove the headers
  itemList.shift();
  sortedItems.shift();
  const numItems = itemList.length;
  spreadsheet.getSheetByName('Item List').clearContents().getRange(1, 1, numItems).setValues(itemList);
  spreadsheet.getSheetByName('Recently Created').clearContents().getRange(1, 1, numItems).setValues(sortedItems);
}

/**
 * This function looks at the UPC database and removes all of the barcodes that are not UPC-A or EAN-13. It also updates the data with the typical Google sheets description string.
 * 
 * @author Jarren Ralf
 */
function updateUPCs()
{
  var sku_upc, item;
  const adagioInventory = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  const itemNum = adagioInventory[0].indexOf('Item #')
  const fullDescription = adagioInventory[0].indexOf('Item List')
  const data = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString()).filter(upc => isUPC_A(upc[0]) || isEAN_13(upc[0])).map(upcs => {
    sku_upc = upcs[1].toUpperCase()
    item = adagioInventory.find(sku => sku[itemNum] === sku_upc)
    return (item != null) ? [upcs[0], item[fullDescription]] : null;
  }).filter(val => val != null).sort(sortUPCsNumerically)

  SpreadsheetApp.getActive().getSheetByName('UPC Database').clearContents().getRange(1, 1, data.length, data[0].length).setNumberFormat('@').setValues(data);
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
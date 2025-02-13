/*
 * Alireza Hakim
 * February 1 2024
 * This program automates the creation of Meeting Minutes for Finance Meetings in LES
 * alirez.hakim@gmail.com for help with the program
 */

//Makes the menu item in Google Sheets
function onOpen() 
{
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Requests');
  menu.addItem('Create meeting minutes from current cell onwards', 'createRequestTable');
  menu.addItem('Create meeting minutes for selected requests', 'multiSelectRequest')
  menu.addToUi();
}//end onOpen()
//Method to create the meeting minutes doc
function createRequestTable() 
{
  //Create variables for template, destination folder, request form sheet, date, new doc, and request sheet array
  const requestSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //MAINTENANCE: Paste file ID of the top half of the meeting minutes template
  const meetingMinutesTemplate = DriveApp.getFileById('1suv4lW9HQE4jKXzeP5xxLjDM4P3qkBY8oIN6PDNrpgI');
  //MAINTENANCE: Paste folder ID of the folder you want the final document to be in
  const destinationFolder = DriveApp.getFolderById('1EX58-6B9d1v2N21endoEC-bEGd2DP5cP');
  const currentDate = new Date().toDateString();
  const docTitle = 'FC Meeting #Num Minutes {' + currentDate.substring(8, 10) + '/' + 
                  currentDate.substring(4, 7) + '/' + currentDate.substring(11) + '}';
  const copy = meetingMinutesTemplate.makeCopy(docTitle, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  const requestTableTemplate = doc.getBody().getTables()[3];
  const rows = requestSheet.getDataRange().getValues();
  var startIndex = requestSheet.getCurrentCell().getRowIndex();
  //Create all the request tables
  rows.forEach(function(row, index) 
  {
    if (startIndex == null) throw new Error("Please select a request row to start from.");
    if (index < startIndex - 1) return;
    if (!row[1]) return;
    const requestTable = requestTableTemplate.copy();
    requestTable.replaceText('{{Request ID}}', row[0]);
    requestTable.replaceText('{{Request By}}', row[1]);
    requestTable.replaceText('{{Subject}}', row[2]);
    requestTable.replaceText('{{Budget Line}}', row[3]);
    requestTable.replaceText('{{Expected Price}}', row[4]);
    requestTable.replaceText('{{Description}}', row[6]);
    doc.getBody().appendTable(requestTable);
  });
  requestTableTemplate.removeFromParent();
  //Appends section for the end of the doc
  //MAINTENANCE: Paste the ID of the bottom half of the template doc
  const endDoc = DocumentApp.openById('11YsYDP94eC-ZwGQLPUeIAY_ujva2z4KHL3khQJYxrrg');
  var numElements = endDoc.getBody().getNumChildren();
  var element;
  var type;
  for (var i = 0; i < numElements; i++)
  {
    element = endDoc.getBody().getChild(i).copy();
    type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH)
    {
      doc.getBody().appendParagraph(element);
    }
    else if (type == DocumentApp.ElementType.TABLE)
    {
      doc.getBody().appendTable(element);
    }
    else if (type == DocumentApp.ElementType.LIST_ITEM)
    {
      doc.getBody().appendListItem(element);
    }
    else 
    {
      throw new Error("Unexpected element while appending end of document: " + type)
    }//end if
  }//end loop
  doc.saveAndClose();
}//end createRequestTable() 
function multiSelectRequest() 
{
  /*//Create variables for template, destination folder, request form sheet, date, new doc, and request sheet array
  const requestSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //MAINTENANCE: Paste file ID of the top half of the meeting minutes template
  const meetingMinutesTemplate = DriveApp.getFileById('1suv4lW9HQE4jKXzeP5xxLjDM4P3qkBY8oIN6PDNrpgI');
  //MAINTENANCE: Paste folder ID of the folder you want the final document to be in
  const destinationFolder = DriveApp.getFolderById('1EX58-6B9d1v2N21endoEC-bEGd2DP5cP');
  const currentDate = new Date().toDateString();
  const docTitle = 'FC Meeting #Num Minutes {' + currentDate.substring(8, 10) + '/' + 
                  currentDate.substring(4, 7) + '/' + currentDate.substring(11) + '}';
  const copy = meetingMinutesTemplate.makeCopy(docTitle, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  const requestTableTemplate = doc.getBody().getTables()[3];
  const rowIndices = [];
  var numRows = 0;
  const rows = requestSheet.getDataRange().getValues();
  var selectedRanges = requestSheet.getActiveRangeList().getRanges();
  var startIndex = requestSheet.getCurrentCell().getRowIndex();
  //Retrieve the row indices of all selected ranges (requests)
  for (let i = 0; i < selectedRanges.length; i++)
  {
    for (let j = 0; j < selectedRanges[i].getNumRows(); j++)
    {
      rowIndices[numRows] = selectedRanges[i].getCell(j, 0).getRowIndex();
      numRows++;
    }
  }
  //Create all the request tables
  rowIndices.forEach(function(rowIndex)
  {
    const requestTable = requestTableTemplate.copy();
    requestTable.replaceText('{{Request ID}}', rows[rowIndex][0]);
    requestTable.replaceText('{{Request By}}', rows[rowIndex][1]);
    requestTable.replaceText('{{Subject}}', rows[rowIndex][2]);
    requestTable.replaceText('{{Budget Line}}', rows[rowIndex][3]);
    requestTable.replaceText('{{Expected Price}}', rows[rowIndex][4]);
    requestTable.replaceText('{{Description}}', rows[rowIndex][6]);
    doc.getBody().appendTable(requestTable);
  })
  requestTableTemplate.removeFromParent();
  //Appends section for the end of the doc
  const endDoc = DocumentApp.openById('11YsYDP94eC-ZwGQLPUeIAY_ujva2z4KHL3khQJYxrrg');
  var numElements = endDoc.getBody().getNumChildren();
  var element;
  var type;
  for (var i = 0; i < numElements; i++)
  {
    element = endDoc.getBody().getChild(i).copy();
    type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH)
    {
      doc.getBody().appendParagraph(element);
    }
    else if (type == DocumentApp.ElementType.TABLE)
    {
      doc.getBody().appendTable(element);
    }
    else if (type == DocumentApp.ElementType.LIST_ITEM)
    {
      doc.getBody().appendListItem(element);
    }
    else 
    {
      throw new Error("Unexpected element while appending end of document: " + type)
    }//end if
  }//end loop
  doc.saveAndClose();*/
}//end createRequestTable() 
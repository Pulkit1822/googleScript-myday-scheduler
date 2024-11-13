/**
 * Yeh file global constants aur functions ko define karti hai jo myday Google Sheets add-on ke liye hain.
 */

const THEME={
  COLORS: {
    BG_GRAY: '#f6f8fa',// Background color gray
    BORDER_GRAY: '#d0d7de',// Border color gray
    HEADER_BG: '#24292f',// Header background color
    TASK_BG: '#fff3cd',// Task background color
    ASSIGNMENT_BG: '#d1ecf1',// Assignment background color
    NOTES_BG: '#eaf7ea',// Notes background color
    STATUS_COMPLETED: '#d4edda',// Completed status color
    STATUS_IN_PROGRESS: '#fff3cd',// In progress status color
    STATUS_NOT_STARTED: '#f8d7da' // Not started status color
  },


  STATUSES: ['⭕ Not Started','🟡 In Progress','✅ Completed'] // Status options
};

/**
 * Yeh function Google Sheets UI mein ek menu create karta hai.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('myday') // 'myday' naam ka menu banayega
    .addItem('Add New Day\'s Schedule','addNewDaySchedule') // Menu mein ek item add karega
    .addToUi(); // Menu ko UI mein add karega
}

/**
 * Yeh function ek naye din ka schedule sheet mein add karta hai.
 */



function addNewDaySchedule() {
  const sheet=SpreadsheetApp.getActiveSheet(); // Active sheet ko get karega
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy"); // Aaj ki date ko format karega

  if (sheet.getLastRow()===0) {
    setupSheet(sheet); // Agar sheet empty hai to setupSheet function call karega
  }

  const startRow=sheet.getLastRow()+1; // Last row ke baad se start karega

  // Date header ko styling ke sath add karega
  const dateRange=sheet.getRange(startRow,1,1,8); // 8 columns ko span karega
  dateRange.merge()
    .setValue(`📅 ${today}`) // Aaj ki date set karega
    .setBackground(THEME.COLORS.BORDER_GRAY) // Background color set karega
    .setFontSize(12) // Font size set karega
    .setFontWeight('bold') // Font weight bold karega
    .setHorizontalAlignment('center') // Text ko center align karega
    .setVerticalAlignment('middle'); // Text ko vertically middle align karega



  // Tasks aur Assignments columns ke headers set karega
  const headers=[
    ['S. No.','Tasks','Deadline','Status','S. No.','Assignments','Deadline','Status']
  ];


  sheet.getRange(startRow+1,1,1,8).setValues(headers)
    .setBackground(THEME.COLORS.BG_GRAY) // Background color set karega
    .setFontWeight('bold') // Font weight bold karega
    .setHorizontalAlignment('center') // Text ko center align karega
    .setVerticalAlignment('middle'); // Text ko vertically middle align karega

  // Tasks aur assignments ke rows add karega
  const rows=[
    [1,'','',THEME.STATUSES[0],1,'','',THEME.STATUSES[0]],
    [2,'','',THEME.STATUSES[0],2,'','',THEME.STATUSES[0]],
    [3,'','',THEME.STATUSES[0],3,'','',THEME.STATUSES[0]],
    [4,'','',THEME.STATUSES[0],4,'','',THEME.STATUSES[0]],
    [5,'','',THEME.STATUSES[0],5,'','',THEME.STATUSES[0]]
  ];
  sheet.getRange(startRow+2,1,rows.length,8).setValues(rows);

  // Borders aur cell sizes ko style karega
  const taskRange=sheet.getRange(startRow+1,1,rows.length+1,8);
  taskRange.setBorder(true,true,true,true,true,true);
  
  // Puri table mein text aur dates ko center align karega
  sheet.getRange(startRow+1,1,rows.length+1,8).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Status columns ke liye data validations add karega (Tasks aur Assignments dono ke liye)
  const taskStatusRange=sheet.getRange(startRow+2,4,rows.length,1);
  const assignmentStatusRange=sheet.getRange(startRow+2,8,rows.length,1);
  const statusValidation=SpreadsheetApp.newDataValidation().requireValueInList(THEME.STATUSES).build();
  
  taskStatusRange.setDataValidation(statusValidation);
  assignmentStatusRange.setDataValidation(statusValidation);

  // Deadline columns ke liye data validations add karega (Tasks aur Assignments dono ke liye)
  const taskDeadlineRange=sheet.getRange(startRow+2,3,rows.length,1);
  const assignmentDeadlineRange=sheet.getRange(startRow+2,7,rows.length,1);
  const dateValidation=SpreadsheetApp.newDataValidation().requireDate().build();
  
  taskDeadlineRange.setDataValidation(dateValidation);
  assignmentDeadlineRange.setDataValidation(dateValidation);

  // Deadline columns ko date format mein set karega
  const dateFormat="dd/MM/yyyy";
  taskDeadlineRange.setNumberFormat(dateFormat);
  assignmentDeadlineRange.setNumberFormat(dateFormat);

  // Status columns ke liye conditional formatting add karega
  addConditionalFormatting(sheet,startRow+2,rows.length);

  // Notes section add karega
  const notesRow=startRow+rows.length+2;
  const notesRange=sheet.getRange(notesRow,1,1,8);
  notesRange.merge()
    .setValue('📝 Notes:')
    .setBackground(THEME.COLORS.NOTES_BG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

/**
 * Sheet ko title aur column widths ke sath setup karega.
 */
function setupSheet(sheet) {
  // Title ko distinctive header style ke sath set karega
  const titleRange=sheet.getRange(1,1,1,8);
  titleRange.merge()
    .setValue('🗓️ myday 🗓️')
    .setBackground(THEME.COLORS.HEADER_BG)
    .setFontColor('white')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Column widths ko readability ke liye set karega
  sheet.setColumnWidths(1,8,150);
  sheet.setColumnWidth(1,50);  // Serial No.
  sheet.setColumnWidth(2,200); // Tasks
  sheet.setColumnWidth(3,100); // Deadline
  sheet.setColumnWidth(4,120); // Status
  sheet.setColumnWidth(5,50);  // Serial No.
  sheet.setColumnWidth(6,200); // Assignments
  sheet.setColumnWidth(7,100); // Deadline
  sheet.setColumnWidth(8,120); // Status
}

/**
 * Sheet mein conditional formatting add karega.
 */
function addConditionalFormatting(sheet,startRow,numRows) {
  const rules=[];
  
  // Status columns (column 4 aur 8) ke liye rules add karega
  [4,8].forEach(col=> {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('✅ Completed')


        .setBackground(THEME.COLORS.STATUS_COMPLETED)
        .setFontColor('#155724')
        .setRanges([sheet.getRange(startRow,col,numRows)])
        .build(),
        


      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('⭕ Not Started')
        .setBackground(THEME.COLORS.STATUS_NOT_STARTED)
        .setFontColor('#721c24')
        .setRanges([sheet.getRange(startRow,col,numRows)])
        .build()
    );
  });
  
  sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(rules));
}

/**
 * Sheet mein onEdit event ko handle karega.
 */
function onEdit(e) {
  const sheet=e.source.getActiveSheet();
  const range=e.range;
  const nextRow=range.getRow()+1;


  const col=range.getColumn();
  
  // Sirf task aur assignment description columns (columns 2 aur 6) ke liye proceed karega
  if ((col===2 || col===6) && range.getValue() !=='') {

    // Agar next row empty hai to usko activate karega
    const nextCell=sheet.getRange(nextRow,col);
    if (nextCell.getValue()==='') {
      nextCell.activate();
    }
  }
}
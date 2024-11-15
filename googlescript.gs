// Theme and Constants - ye saare colors aur basic settings hai 
const THEME={
  COLORS: {
    BG_GRAY: '#f6f8fa',    // background ke liye light gray
    BORDER_GRAY: '#d0d7de', // border ke liye thoda dark gray
    HEADER_BG: '#24292f',   // header ka background ekdum black jaisa
    TASK_BG: '#fff3cd',     // normal tasks ke liye halka yellow
    ASSIGNMENT_BG: '#d1ecf1',// assignments ke liye light blue
    NOTES_BG: '#eaf7ea',    // notes section ke liye mint green types
    STATUS_COMPLETED: '#d4edda',   // completed waale tasks ke liye green
    STATUS_IN_PROGRESS: '#fff3cd', // jo abhi chal rahe hai unke liye yellow
    STATUS_NOT_STARTED: '#f8d7da'   // jo start hi nahi huye unke liye red
  },
  STATUSES: ['⭕ Not Started','🟡 In Progress','✅ Completed'] // status ke teen options bas
};

// Calendar Settings - Google Calendar ke liye settings
const CALENDAR_SETTINGS={
  REMINDER_TIMES: [
    24 * 60,// ek din pehle reminder
    60,     // 1 ghanta pehle
    10       // last 10 minute mein bhi ek reminder
  ],
  EVENT_COLORS: {

    TASK: '10',    // normal tasks green color mein dikhenge
    ASSIGNMENT: '9'  // assignments purple mein,easy identify ke liye
  }
};

// bhai calendar access ke liye permission mangne wala function
function requestCalendarPermission() {

  const calendar=CalendarApp.getDefaultCalendar();
  Logger.log(calendar.getName());  // ye permission prompt karega
}

// jab sheet khulegi tab ye menu add hoga - ekdum simple
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('myday')
    .addItem('Add New Day\'s Schedule','addNewDaySchedule')
    .addItem('Sync All Events to Calendar','syncAllEventsToCalendar')
    .addItem('Initialize Calendar Access','requestCalendarPermission')
    .addToUi();
}

// OAuth token nikalne ke liye - authentication ke liye jaruri hai ye
function getOAuthToken() {

  return ScriptApp.getOAuthToken();
}

// Check if event already exists for that day
function eventExistsForDay(taskTitle,deadline,isAssignment=false) {
  try {


    const calendar=CalendarApp.getDefaultCalendar();
    const events=calendar.getEventsForDay(new Date(deadline));
    const searchTitle=`${isAssignment ? '📚 Assignment' : '✔️ Task'}: ${taskTitle}`;
    
    return events.some(event=> event.getTitle()===searchTitle);
  } catch (error) {


    console.error('error in event:',error);
    return false;
  }
}

// Modified createCalendarEvent function with duplicate check
function createCalendarEvent(taskTitle,deadline,isAssignment=false) {
  try {

    // Pehle check karo ki event already exists to nahi
    if (eventExistsForDay(taskTitle,deadline,isAssignment)) {
      throw new Error('Event already exists for this day');
    }

    const calendar=CalendarApp.getDefaultCalendar();
    const eventTitle=`${isAssignment ? '📚 Assignment' : '✔️ Task'}: ${taskTitle}`;
    
    const event=calendar.createAllDayEvent(
      eventTitle,
      new Date(deadline)
    );
    
    CALENDAR_SETTINGS.REMINDER_TIMES.forEach(time=> {
      event.addPopupReminder(time);
      event.addEmailReminder(time);
    });
    
    event.setColor(isAssignment ? 
      CALENDAR_SETTINGS.EVENT_COLORS.ASSIGNMENT : 
      CALENDAR_SETTINGS.EVENT_COLORS.TASK
    );
    
    return event.getId();
  } catch (error) {
    console.error('Error occured while syncing:',error);
    throw error;
  }
}

// koi event delete karna ho to ye function use karo
function removeCalendarEvent(date,taskTitle) {


  const calendar=CalendarApp.getDefaultCalendar();
  const events=calendar.getEventsForDay(new Date(date));
  
  // jo title match karega wo event delete ho jayega
  events.forEach(event=> {
    if (event.getTitle().includes(taskTitle)) {

      event.deleteEvent();
    }
  });
}

// Modified syncAllEventsToCalendar function with better error handling
function syncAllEventsToCalendar() {
  const sheet=SpreadsheetApp.getActiveSheet();
  const data=sheet.getDataRange().getValues();

  let currentRow=2; // header ke baad se start karo
  
  data.forEach((row,index)=> {
    // pehli row aur khaali rows ko skip karo
    if (index < 1 || !row[0]) return;
    
    // tasks ko process karo
    if (row[2]) { // agar deadline set hai to
      try {
        if (!eventExistsForDay(row[1],row[2],false)) {

          createCalendarEvent(row[1],row[2],false);
          sheet.getRange(currentRow,3).setNote('📅 Added in Calendar ');
        } else {

          sheet.getRange(currentRow,3).setNote('ℹ️ Event already exists');
        }
      } catch (error) {
        sheet.getRange(currentRow,3).setNote('❌ Not Synced: '+error.message);
      }
    }
    
    // assignments ko process karo
    if (row[6]) { // agar deadline set hai to
      try {

        if (!eventExistsForDay(row[5],row[6],true)) {

          createCalendarEvent(row[5],row[6],true);
          sheet.getRange(currentRow,7).setNote('📅 Added in Calendar');
        } else {

          sheet.getRange(currentRow,7).setNote('ℹ️ Event already exists');
        }
      } catch (error) {

        sheet.getRange(currentRow,7).setNote('❌ Not Synced: '+error.message);
      }
    }
    
    currentRow++;
  });
}

// sheet ko setup karne ka function - ekdum basic structure banayega
function setupSheet(sheet) {


  // sabse upar title lagao,ekdum mast style ke saath
  const titleRange=sheet.getRange(1,1,1,8);
  titleRange.merge()
    .setValue('🗓️ myday 🗓️')
    .setBackground(THEME.COLORS.HEADER_BG)
    .setFontColor('white')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // columns ki width set karo - sab kuch readable rakho
  sheet.setColumnWidth(1,50);   // S.No. ke liye choti si
  sheet.setColumnWidth(2,200);  // Tasks ke liye badi
  sheet.setColumnWidth(3,100);  // Deadline medium
  sheet.setColumnWidth(4,120);  // Status ke liye thodi badi
  sheet.setColumnWidth(5,50);   // Assignments ka S.No.
  sheet.setColumnWidth(6,200);  // Assignment description ke liye badi
  sheet.setColumnWidth(7,100);  // Assignment deadline
  sheet.setColumnWidth(8,120);  // Assignment status
}

// conditional formatting add karne ka system - status ke hisaab se color change hoga
function addConditionalFormatting(sheet,startRow,numRows) {
  const rules=[];
  
  // Status columns ke liye formatting rules
  [4,8].forEach(col=> {
    rules.push(
      // Complete ho gaya to green
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('✅ Completed')
        .setBackground(THEME.COLORS.STATUS_COMPLETED)
        .setFontColor('#155724')
        .setRanges([sheet.getRange(startRow,col,numRows)])
        .build(),
      
      // In progress hai to yellow
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('🟡 In Progress')
        .setBackground(THEME.COLORS.STATUS_IN_PROGRESS)
        .setFontColor('#856404')
        .setRanges([sheet.getRange(startRow,col,numRows)])
        .build(),
      
      // Start nahi hua to red
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

// naye din ka schedule add karne ka system
function addNewDaySchedule() {
  const sheet=SpreadsheetApp.getActiveSheet();
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy");

  // agar sheet bilkul khaali hai to pehle setup karo
  if (sheet.getLastRow()===0) {
    setupSheet(sheet);
  }

  const startRow=sheet.getLastRow()+1;

  // today's date's header
  const dateRange=sheet.getRange(startRow,1,1,8);
  dateRange.merge()
    .setValue(`📅 ${today}`)
    .setBackground(THEME.COLORS.BORDER_GRAY)
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // columns ke headers
  const headers=[
    ['S. No.','Tasks','Deadline','Status','S. No.','Assignments','Deadline','Status']
  ];
  sheet.getRange(startRow+1,1,1,8).setValues(headers)
    .setBackground(THEME.COLORS.BG_GRAY)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // 5 blank rows  - by default
  const rows=Array(5).fill().map((_,i)=> [
    i+1,'','',THEME.STATUSES[0],
    i+1,'','',THEME.STATUSES[0]
  ]);
  
  const dataRange=sheet.getRange(startRow+2,1,rows.length,8);
  dataRange.setValues(rows);

  // cells ka thoda styling 
  const taskRange=sheet.getRange(startRow+1,1,rows.length+1,8);
  taskRange.setBorder(true,true,true,true,true,true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // status ke liye dropdown validation
  const statusValidation=SpreadsheetApp.newDataValidation()
    .requireValueInList(THEME.STATUSES)
    .build();
  
  [4,8].forEach(col=> {

    sheet.getRange(startRow+2,col,rows.length,1)
      .setDataValidation(statusValidation);
  });

  // date columns ke liye validation
  const dateValidation=SpreadsheetApp.newDataValidation()
    .requireDate()
    .build();
  
  [3,7].forEach(col=> {
    const deadlineRange=sheet.getRange(startRow+2,col,rows.length,1);
    deadlineRange.setDataValidation(dateValidation)
      .setNumberFormat("dd/MM/yyyy");
  });

  // conditional formatting bhi add karo
  addConditionalFormatting(sheet,startRow+2,rows.length);

  // last mein notes ka section
  const notesRow=startRow+rows.length+2;
  sheet.getRange(notesRow,1,1,8).merge()
    .setValue('📝 Notes:')
    .setBackground(THEME.COLORS.NOTES_BG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

// jab kuch edit ho tab ye function chalega
function onEdit(e) {
  const sheet=e.source.getActiveSheet();
  const range=e.range;
  const row=range.getRow();
  const col=range.getColumn();
  const value=range.getValue();
  
  // task ya assignment add karne ke baad automatically next cell pe focus
  if ((col===2 || col===6) && value !=='') {
    const nextCell=sheet.getRange(row+1,col);

    if (nextCell.getValue()==='') {
      nextCell.activate();
    }
  }
  
  // deadline add karne pe calendar mein event add ho jayega
  if ((col===3 || col===7) && value !=='') {
    try {
      const isAssignment=(col===7);
      const taskCol=isAssignment ? 6 : 2;
      const taskTitle=sheet.getRange(row,taskCol).getValue();
      
      if (taskTitle) {
        if (!eventExistsForDay(taskTitle,value,isAssignment)) {

          createCalendarEvent(taskTitle,value,isAssignment);
          
          // status update karo agar jarurat hai to
          const statusCell=sheet.getRange(row,col+1);
          if (statusCell.getValue()===THEME.STATUSES[0]) {
            statusCell.setValue(THEME.STATUSES[1]);
          }
          
          range.setNote('📅 Added in Calendar');
        } else {
          range.setNote('ℹ️ Event already exists');
        }
      } else {

        range.setNote('⚠️ Add title for task');
      }
    } catch (error) {

      range.setNote('❌ Error Occured: '+error.message);
    }
  }
  
  // deadline hatane pe calendar se bhi event hat jayega
  if ((col===3 || col===7) && e.oldValue && !value) {
    try {

      const taskCol=(col===7) ? 6 : 2;
      const taskTitle=sheet.getRange(row,taskCol).getValue();
      removeCalendarEvent(e.oldValue,taskTitle);
      range.setNote('Deleted from Calender');
    } catch (error) {

      range.setNote('❌ Error Occured: '+error.message);
    }
  }
}

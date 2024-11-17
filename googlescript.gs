const THEME = {
  COLORS: {
    BG_GRAY: '#fdf2f8',        // Soft pink-white
    BORDER_GRAY: '#000000',    // Black borders for pop
    HEADER_BG: '#ff6b6b',      // Coral pink - energetic & fun
    DATE_DIVIDER_BG: '#ffd93d', // Sunshine yellow
    TASK_BG: {
      DEFAULT: '#e6f7ff',      // Bright baby blue
      ALTERNATE: '#e6f7ff'     // Whisper blue (light-tinted white)
    },
    ASSIGNMENT_BG: {
      DEFAULT: '#ffe4ee',      // Bright cotton candy pink
      ALTERNATE: '#ffe4ee'     // Whisper pink (light-tinted white)
    },
    NOTES_BG: '#c4faf8',       // Mint turquoise
    STATUS: {
      COMPLETED: {
        BG: '#86efac',         // Success green
        TEXT: '#052e16'        // Dark green text
      },
      IN_PROGRESS: {
        BG: '#fcd34d',         // Warm yellow
        TEXT: '#451a03'        // Dark brown text
      },
      NOT_STARTED: {
        BG: '#fca5a5',         // Soft red
        TEXT: '#450a0a'        // Dark red text
      }
    },
    BORDERS: {
      TASK: '#000000',         // Black borders
      ASSIGNMENT: '#000000',   // Black borders
      HEADER: '#000000'        // Black borders
    }
  },
  STATUSES: ['⭕ Not Started', '🟡 In Progress', '✅ Completed'],
  STYLING: {
    BORDER_RADIUS: '12px',     // Rounded for fun vibes
    SHADOW: '0 4px 12px rgba(0,0,0,0.08)', // Playful shadow
    HEADER_FONT: 'Quicksand, sans-serif',  // Round, friendly font
    CONTENT_FONT: 'Quicksand, sans-serif'  // Matching friendly font
  }
};

// THEME constant defines the color scheme and styling for the sheet

function requestCalendarPermission() {
  const calendar = CalendarApp.getDefaultCalendar();
  Logger.log(calendar.getName());
}

// This function requests permission to access the default calendar

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('myday')
    .addItem('Add New Day\'s Schedule', 'addNewDaySchedule')
    .addItem('Sync All Events to Calendar', 'syncAllEventsToCalendar')
    .addItem('Initialize Calendar Access', 'requestCalendarPermission')
    .addToUi();
}

// This function creates a custom menu in the Google Sheets UI when the spreadsheet is opened

function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

// This function returns the OAuth token for the script

function eventExistsForDay(taskTitle, deadline, isAssignment = false) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEventsForDay(new Date(deadline));
    const searchTitle = `${isAssignment ? '📚 Assignment' : '✔️ Task'}: ${taskTitle}`;
    return events.some(event => event.getTitle() === searchTitle);
  } catch (error) {
    console.error('error in event:', error);
    return false;
  }
}

// This function checks if an event with the given title and deadline already exists in the calendar

function createCalendarEvent(taskTitle, deadline, isAssignment = false) {
  try {
    if (eventExistsForDay(taskTitle, deadline, isAssignment)) {
      throw new Error('Event already exists for this day');
    }

    const calendar = CalendarApp.getDefaultCalendar();
    const eventTitle = `${isAssignment ? '📚 Assignment' : '✔️ Task'}: ${taskTitle}`;
    
    const event = calendar.createAllDayEvent(
      eventTitle,
      new Date(deadline)
    );
    
    CALENDAR_SETTINGS.REMINDER_TIMES.forEach(time => {
      event.addPopupReminder(time);
      event.addEmailReminder(time);
    });
    
    event.setColor(isAssignment ? 
      CALENDAR_SETTINGS.EVENT_COLORS.ASSIGNMENT : 
      CALENDAR_SETTINGS.EVENT_COLORS.TASK
    );
    
    return event.getId();
  } catch (error) {
    console.error('Error occured while syncing:', error);
    throw error;
  }
}

// This function creates a new calendar event for the given task or assignment

function removeCalendarEvent(date, taskTitle) {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEventsForDay(new Date(date));
  events.forEach(event => {
    if (event.getTitle().includes(taskTitle)) {
      event.deleteEvent();
    }
  });
}

// This function removes a calendar event with the given title and date

function syncAllEventsToCalendar() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let currentRow = 2;
  
  data.forEach((row, index) => {
    if (index < 1 || !row[0]) return;
    
    if (row[2]) {
      try {
        if (!eventExistsForDay(row[1], row[2], false)) {
          createCalendarEvent(row[1], row[2], false);
          sheet.getRange(currentRow, 3).setNote('📅 Added in Calendar ');
        } else {
          sheet.getRange(currentRow, 3).setNote('ℹ️ Event already exists');
        }
      } catch (error) {
        sheet.getRange(currentRow, 3).setNote('❌ Not Synced: ' + error.message);
      }
    }
    
    if (row[6]) {
      try {
        if (!eventExistsForDay(row[5], row[6], true)) {
          createCalendarEvent(row[5], row[6], true);
          sheet.getRange(currentRow, 7).setNote('📅 Added in Calendar');
        } else {
          sheet.getRange(currentRow, 7).setNote('ℹ️ Event already exists');
        }
      } catch (error) {
        sheet.getRange(currentRow, 7).setNote('❌ Not Synced: ' + error.message);
      }
    }
    currentRow++;
  });
}

// This function syncs all events from the Google Sheet to the Google Calendar

function setupSheet(sheet) {
  const titleRange = sheet.getRange(1, 1, 1, 8);
  titleRange.merge()
    .setValue('✨ MyDay Planner ✨')
    .setBackground(THEME.COLORS.HEADER_BG)
    .setFontColor('white')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  titleRange.setBorder(
    true, true, true, true, false, false,
    THEME.COLORS.BORDERS.HEADER,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 250);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 140);
}

// This function sets up the initial layout and styling for the sheet

function addConditionalFormatting(sheet, startRow, numRows) {
  const rules = [];
  
  [4, 8].forEach(col => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('✅ Completed')
        .setBackground(THEME.COLORS.STATUS.COMPLETED.BG)
        .setFontColor(THEME.COLORS.STATUS.COMPLETED.TEXT)
        .setRanges([sheet.getRange(startRow, col, numRows)])
        .build(),
      
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('🟡 In Progress')
        .setBackground(THEME.COLORS.STATUS.IN_PROGRESS.BG)
        .setFontColor(THEME.COLORS.STATUS.IN_PROGRESS.TEXT)
        .setRanges([sheet.getRange(startRow, col, numRows)])
        .build(),
      
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('⭕ Not Started')
        .setBackground(THEME.COLORS.STATUS.NOT_STARTED.BG)
        .setFontColor(THEME.COLORS.STATUS.NOT_STARTED.TEXT)
        .setRanges([sheet.getRange(startRow, col, numRows)])
        .build()
    );
  });
  
  sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(rules));
}

// This function adds conditional formatting rules to the sheet based on the task status

function addNewDaySchedule() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  if (sheet.getLastRow() === 0) {
    setupSheet(sheet);
  }

  const startRow = sheet.getLastRow() + 1;

  const dateRange = sheet.getRange(startRow, 1, 1, 8);
  dateRange.merge()
    .setValue(`📅 ${today}`)
    .setBackground(THEME.COLORS.DATE_DIVIDER_BG)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  dateRange.setBorder(
    true, true, true, true, false, false,
    THEME.COLORS.BORDER_GRAY,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  const headers = [['№', 'Tasks', 'Deadline', 'Status', '№', 'Assignments', 'Deadline', 'Status']];
  const headerRange = sheet.getRange(startRow + 1, 1, 1, 8);
  headerRange.setValues(headers)
    .setBackground(THEME.COLORS.BG_GRAY)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  headerRange.setBorder(
    true, true, true, true, true, true,
    THEME.COLORS.BORDER_GRAY,
    SpreadsheetApp.BorderStyle.SOLID
  );

  const rows = Array(5).fill().map((_, i) => [
    i + 1, '', '', THEME.STATUSES[0],
    i + 1, '', '', THEME.STATUSES[0]
  ]);
  
  const dataRange = sheet.getRange(startRow + 2, 1, rows.length, 8);
  dataRange.setValues(rows);

  for (let i = 0; i < rows.length; i++) {
    const rowRange = sheet.getRange(startRow + 2 + i, 1, 1, 4);
    const assignmentRange = sheet.getRange(startRow + 2 + i, 5, 1, 4);
    
    rowRange.setBackground(i % 2 === 0 ? THEME.COLORS.TASK_BG.DEFAULT : THEME.COLORS.TASK_BG.ALTERNATE);
    assignmentRange.setBackground(i % 2 === 0 ? THEME.COLORS.ASSIGNMENT_BG.DEFAULT : THEME.COLORS.ASSIGNMENT_BG.ALTERNATE);
  }

  const taskRange = sheet.getRange(startRow + 1, 1, rows.length + 1, 8);
  taskRange.setBorder(
    true, true, true, true, true, true,
    THEME.COLORS.BORDER_GRAY,
    SpreadsheetApp.BorderStyle.SOLID
  )
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(THEME.STATUSES)
    .build();
  
  [4, 8].forEach(col => {
    sheet.getRange(startRow + 2, col, rows.length, 1)
      .setDataValidation(statusValidation);
  });

  const dateValidation = SpreadsheetApp.newDataValidation()
    .requireDate()
    .build();
  
  [3, 7].forEach(col => {
    const deadlineRange = sheet.getRange(startRow + 2, col, rows.length, 1);
    deadlineRange.setDataValidation(dateValidation)
      .setNumberFormat("dd/MM/yyyy");
  });

  addConditionalFormatting(sheet, startRow + 2, rows.length);

  const notesRow = startRow + rows.length + 2;
  const notesRange = sheet.getRange(notesRow, 1, 1, 8);
  notesRange.merge()
    .setValue('📝 Notes & Reflections:')
    .setBackground(THEME.COLORS.NOTES_BG)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(
      true, true, true, true, false, false,
      THEME.COLORS.BORDERS.TASK,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
}

// This function adds a new day's schedule to the sheet with predefined tasks and assignments

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = range.getValue();
  
  if ((col === 2 || col === 6) && value !== '') {
    const nextCell = sheet.getRange(row + 1, col);
    if (nextCell.getValue() === '') {
      nextCell.activate();
    }
  }
  
  if ((col === 3 || col === 7) && value !== '') {
    try {
      const isAssignment = (col === 7);
      const taskCol = isAssignment ? 6 : 2;
      const taskTitle = sheet.getRange(row, taskCol).getValue();
      
      if (taskTitle) {
        if (!eventExistsForDay(taskTitle, value, isAssignment)) {
          createCalendarEvent(taskTitle, value, isAssignment);
          
          const statusCell = sheet.getRange(row, col + 1);
          if (statusCell.getValue() === THEME.STATUSES[0]) {
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
      range.setNote('❌ Error Occured: ' + error.message);
    }
  }
  
  if ((col === 3 || col === 7) && e.oldValue && !value) {
    try {
      const taskCol = (col === 7) ? 6 : 2;
      const taskTitle = sheet.getRange(row, taskCol).getValue();
      removeCalendarEvent(e.oldValue, taskTitle);
      range.setNote('Deleted from Calender');
    } catch (error) {
      range.setNote('❌ Error Occured: ' + error.message);
    }
  }
}

// This function handles edits made to the sheet and updates the calendar events accordingly

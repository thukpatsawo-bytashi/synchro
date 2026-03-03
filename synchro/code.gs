
function createYearSheets() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete all old month sheets created by this script
  let allSheets = ss.getSheets();
  allSheets.forEach(s => {
    let name = s.getName();
    if (/^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/.test(name)) {
      ss.deleteSheet(s);
    }
  });

  // July (start) → June (end next year)
  let startYear = new Date().getFullYear();
  let startMonth = 6; // July (0-based)
  let months = 12;

  for (let i = 0; i < months; i++) {
    let monthIndex = (startMonth + i) % 12;
    let year = startYear + Math.floor((startMonth + i) / 12);

    let monthName = new Date(year, monthIndex).toLocaleString("default", { month: "long" });
    let sheetName = monthName + " " + year;

    // Create fresh sheet
    let sheet = ss.insertSheet(sheetName);

    
    sheet.getRange("A1").setValue(new Date(year, monthIndex, 1));

    
    sheet.getRange("J1").setValue("Holiday List (Enter dates)");

    
    sheet.getRange("K1").setValue("No of Business Days");
    sheet.getRange("L1").setValue("No of Holidays/Weekends");

    // Draw  calendar
    drawCalendar(sheet, year, monthIndex);
  }
}

function drawCalendar(sheet, year, monthIndex) {
  let days = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

  // Clear calendar area
  sheet.getRange("A3:H300").clear();

  let row = 3;
  sheet.getRange(row, 1, 1, 7).setValues([days]);
  row++;

  let firstDay = new Date(year, monthIndex, 1);
  let lastDay = new Date(year, monthIndex + 1, 0);

  let col = firstDay.getDay() + 1;
  let currentRow = row;

  for (let d = 1; d <= lastDay.getDate(); d++) {
    let dateCell = sheet.getRange(currentRow, col);
    dateCell.setValue(d);

    col++;
    if (col > 7) {
      col = 1;
      currentRow += 3;
    }
  }

 
  refreshHolidays(sheet, year, monthIndex);
}

function refreshHolidays(sheet, year, monthIndex) {
  
  let holidayValues = sheet.getRange("J2:J").getValues().flat().filter(String);
  let holidays = holidayValues.map(d => new Date(d).setHours(0,0,0,0));

  let firstDay = new Date(year, monthIndex, 1);
  let lastDay = new Date(year, monthIndex + 1, 0);

  let row = 4; // Calendar starts here
  let col = firstDay.getDay() + 1;
  let currentRow = row;
  let businessDay = 0;

  for (let d = 1; d <= lastDay.getDate(); d++) {
    let currentDate = new Date(year, monthIndex, d);
    let weekday = currentDate.getDay();
    let dateKey = currentDate.setHours(0,0,0,0);

    let dateCell = sheet.getRange(currentRow, col);
    let labelCell = sheet.getRange(currentRow + 1, col);

    
    dateCell.setBackground(null);
    labelCell.clearContent();

    if (weekday === 0 || weekday === 6 || holidays.includes(dateKey)) {
      dateCell.setBackground("#f4cccc"); // holiday/weekend
    } else {
      dateCell.setBackground("#f3f3f3"); // business day
      labelCell.setValue("Business Day " + (businessDay + 1));
      businessDay++;
    }

    col++;
    if (col > 7) {
      col = 1;
      currentRow += 3;
    }
  }

  // Update totals
  let totalDays = lastDay.getDate();
  sheet.getRange("K2").setValue(businessDay);
  sheet.getRange("L2").setValue(totalDays - businessDay);
}

function onEdit(e) {
  if (!e) return; 
  let sheet = e.source.getActiveSheet();
  let cell = e.range;

  
  if (cell.getColumn() === 10 && cell.getRow() > 1) {
    let header = sheet.getRange("A1").getValue();
    if (Object.prototype.toString.call(header) === "[object Date]") {
      let year = header.getFullYear();
      let monthIndex = header.getMonth();
      refreshHolidays(sheet, year, monthIndex);
    }
  }
}



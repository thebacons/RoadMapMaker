/* 
Developed by Colin  Bacon 
04-01.2023 - Added the function to add days without distorting the orignal fucntion -> https://codingbeautydev.com/blog/javascript-add-days-to-date/
This enables to dynmically enter any start date. The duration is hardcoded to 400 days but this could be variable as well on line 114- 

1. Set date on line 85 and run on a clean canvas without any merged cells 

Ideas TO DO
// 
1. Offset using something like this dayRange=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1,canvasOffset_j,dimensions[0],dimensions[1]);
2. Lines for milestones and other Gantt elements
4. Complete the size of the sheet and format so it fits on a powerpoint as a Plan On A Page
5. 

*/

/* In this version I'm trying to get the weeks to start on Monday rather than sunday as follows,
Date.prototype.getWeek = function() {
    var onejan = new Date(this.getFullYear(),0,1);
    return Math.ceil((((this - onejan) / 86400000) + onejan.getDay()+1)/7);
} 

var now = new Date();
Logger.log(now.getWeek());
*/
//Set Global variables
//const RGB1 = "#0e316e"; //Dark Blue cell background
//const RGB2 = "#1e7ec7"; //Light Blue cell background
//const RGB3 = "#7209db"; //Pink for weekends

function paintDateRangeDate(){
  //Say Hello in a popup box
  //let myName=Browser.msgBox("Hello");
  //const ui = SpreadsheetApp.getUi();
  //ui.createMenu('My Custom Menu')
  //   .addItem('Say Hello', 'helloWorld')
  //   .addToUi();
   
  
  let toggleDay = true;
  var toggleYear= true;
  let dayBackgroundColour = RGB1; //Start value which toggles
  let weekBackgroundColour = RGB1;//Start value which toggles
  let monthBackgroundColour = RGB1;//Start value which toggles
  let yearBackgroundColour = RGB1;//Start value which toggles
  let rangeBackgroundColour = RGB1;
  let toggleColour = true;
  let toggleYearColour = true;
  let toggleMonthColour = true;
  let toggleWeekColour = true;
   
   
   //Variables
  let i = 1; //top row
  let j=1; //canvas j . This is the starting point for spreadsheet
  
  let dateArray= [];  //Declare empty multidimensional array
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roadMapSheet = ss.getSheets()[0];
  let weekStartAddress= "A3"; //Intial value for 1st loop cycle
  let monthStartAddress="A3"; //Intial value for 1st loop cycle
  let yearStartAddress= "A3"; //Intial value for 1st loop cycle
  let weekEndRange= "";
  let currentYearRange="";
  let currentMonthRange="";
  let currentWeekRange="";
  
  let currentDay=0 // 0 is today's date and each loop is increase currentDay++
  
  
  /*Experimental AREA
    to position the canvasa anywhere on the spreadshhet 
  */
  
    let canvasOffset_j=1; //must be >0 
  
  
  
  // the Utiliites.formatDate give an incorrect ISO week number using an alternative function for week
  //let currentWeekNumber = Utilities.formatDate(new Date(Date.now() + (1000*60*60*24*currentDay)), "GMT", "w"); //To avoid going into  week trigger before making 
  //let currentWeekNumber = getISOWeekNumber(date); 

  //const date = new Date();

  let inputDate = "2022/01/01";
  
  const date = new Date(inputDate);
  console.log("date= " + date);
  let startDate=date.setDate(date.getDate());
  console.log("startDate = "+ startDate); 
  //let addDate= date.setDate(date.getDate());
  //let addDate= date.setDate(date.getDate() + currentDay);
  let startDay=date.getDay();
  let startMonth=date.getMonth();
  let startYear=date.getFullYear();
  
  //let currentWeekNumber = getISOWeekNumber(date);
  let currentWeekNumber = getISOWeekNumber(new Date(addDaysCopy(date, currentDay)))
  console.log('It\'s currently week ' + currentWeekNumber[1] + ' of ' + currentWeekNumber[0]);
  let oldWeekNumber = currentWeekNumber;
  
  //let oldMonth = Utilities.formatDate(new Date(Date.now() + (1000*60*60*24*currentDay)), "GMT", "MMMM");
  //let oldYear = Utilities.formatDate(new Date(Date.now() + (1000*60*60*24*currentDay)), "GMT", "y"); 
  
  //let oldMonth =  Utilities.formatDate(new Date(addDate), "GMT", "MMMM");
  //let oldYear  = Utilities.formatDate(new Date(addDate), "GMT", "y"); 
  let oldMonth = Utilities.formatDate(new Date(addDaysCopy(date, currentDay)), "Europe/Berlin", "MMMM");
  let oldYear = Utilities.formatDate(new Date(addDaysCopy(date, currentDay)), "Europe/Berlin", "y");


  let weekCellRangeCount = j;//the start cell for the week range
  
 //****************************Duration*********************** */
  let numberOfDaysInRange=400; //This determine the number of days to include in the calendar > 1 year
  
  
  //Utilities.formatDate(new Date.now()+ 1000*60*60*24, "GMT+1", "dd/MM/yyyy")
  for (let j=1;j<numberOfDaysInRange;j++){
    //Get the address of the new active cell
    //Add one the cells


     const date = new Date(inputDate);

    let currentCell=roadMapSheet.getRange(3,j);
    currentCell.activate();
    let rowIndex=currentCell.getRowIndex();
    let columnIndex=currentCell.getColumnIndex();
     console.log("j | rowIndex  | columnIndex");
    console.log(j,rowIndex,columnIndex);
    
    let currentCellAddress=roadMapSheet.getActiveCell().getA1Notation();
    console.log("currentCellAddress = " + currentCellAddress);


    //Determine the date data for the array
    //var currentYear =  Utilities.formatDate(new Date(Date("01.12.2021") + (1000*60*60*24*currentDay)), "Europe/Berlin", "y"); 
    //var currentMonth =  Utilities.formatDate(new Date(Date("01.12.2021") + (1000*60*60*24*currentDay)), "Europe/Berlin", "MMMM");
    //var currentDate = Utilities.formatDate(new Date(Date("01.12.2021") + (1000*60*60*24*currentDay)),"Europe/Berlin", "dd/MM/yyyy");
    

    
    //Determine the new date after adding the currentDay++ to the startDate    
     
    
    let currentYear = Utilities.formatDate(new Date(addDaysCopy(date, currentDay)), "Europe/Berlin", "y");
    let currentMonth = Utilities.formatDate(new Date(addDaysCopy(date, currentDay)), "Europe/Berlin", "MMMM");
    let currentDate = Utilities.formatDate(new Date(addDaysCopy(date, currentDay)), "Europe/Berlin", "dd/MM/yyyy");
     
    //let currentYear =  Utilities.formatDate(new Date(addDate), "Europe/Berlin", "y");
    //let currentMonth  = Utilities.formatDate(new Date(addDate), "Europe/Berlin", "MMMM");
    //let currentDate  = Utilities.formatDate(new Date(addDate), "Europe/Berlin", "dd/MM/yyyy");
    console.log("currentDate = " + currentDate);
    




    /* The getISOWeekNumber function is used to determine week number as there is an error with the 
       weeks in the Utilities which gives sunday as the start of the week which is incorrect according to the ISO standard
    */
     

   
    let thisWeek=new Date(addDaysCopy(date, currentDay));
    let currentWeekNumber = getISOWeekNumber(thisWeek);
    
    console.log('It\'s currently week ' + currentWeekNumber[1] + ' of ' + currentWeekNumber[0]);
     var currentDayName = Utilities.formatDate(new Date(addDaysCopy(date, currentDay)), "Europe/Berlin", "E");
    //var currentDayName =  Utilities.formatDate(new Date(Date.now() + (1000*60*60*24*currentDay)), "Europe/Berlin", "E"); 
        
    /******************************************************************* 
                       Year Change Trigger
    ********************************************************************/

    if (currentYear != oldYear){
      console.log ("CHANGE OF YEAR Old/New =" + oldYear, currentYear )
      let yearRangeEndDate= dateArray[currentDay-1][7];
      console.log("yearRangeEndDate= " + yearRangeEndDate)
      currentYearRange=(yearStartAddress +":" + yearRangeEndDate);
      console.log("currentYearRange = " + currentYearRange);

      //Format the YEAR row cell background
      yearBackgroundColour= toggleBackgroundColour(toggleYearColour)
        
    //format the Year
      yearRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentYearRange).offset(-2,0);;
      formatRange(yearRange,14,0,yearBackgroundColour,true,true);//fontSize,Text Rotation, BackgroundColour


    //Store the start values for the next range
    yearStartAddress = currentCellAddress;
    toggleYearColour = !toggleYearColour; //toggle the background colour for the next range
    oldYear = currentYear;
    }


    /******************************************************************* 
                       Month Change Trigger
    ********************************************************************/
    //Determie the month change   
    if (currentMonth != oldMonth){
        //Store the new month value

        console.log ("CHANGE OF MONTH Old/New =" + oldMonth, currentMonth );
        let monthRangeEndDate= dateArray[currentDay-1][7];
        console.log("monthRangeEndDate= " + monthRangeEndDate)
        currentMonthRange=(monthStartAddress +":" + monthRangeEndDate);
        console.log(currentMonthRange);

        //Format the YEAR row cell background
        monthBackgroundColour = toggleBackgroundColour(toggleMonthColour)
        
        //format the month

        monthRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentMonthRange).offset(-1,0);
        formatRange(monthRange,12,0,monthBackgroundColour,true,true);//fontSize,Text Rotation, BackgroundColour
            
        //Store the start values for the next range
        monthStartAddress = currentCellAddress;
        toggleMonthColour = !toggleMonthColour; //toggle the background colour for the next range
        oldMonth = currentMonth;
    }


    /******************************************************************* 
                         Week Change Trigger
    ********************************************************************/
    //Determine week, date and day ranges 
    console.log("WEEK NR = " + currentWeekNumber, oldWeekNumber);
    //console.log(dateArray);
    
    
    if (currentWeekNumber[1] != oldWeekNumber[1]){
      console.log("switched weeks");
      //a trigger happened so read the previous cell's address (which is current j-1) stored in the array
      //and use that value as the rangeEndDate
      let weekRangeEndDate= dateArray[currentDay-1][7];
      console.log("rangeEndDate= " + weekRangeEndDate)
      currentWeekRange=(weekStartAddress +":" + weekRangeEndDate);
      console.log(currentWeekRange);


      //Format and toggle the YEAR row cell background
 
      weekBackgroundColour=toggleBackgroundColour(toggleWeekColour)

    
      //format the week row 3 
      weekRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange);
      formatRange(weekRange,9,0,weekBackgroundColour,true,true);//fontSize,Text Rotation, BackgroundColour
    

      //format the date range row 4 with  text orientation as 90 degrees
      dateRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange).offset(1,0);
      formatRange(dateRange,8,90,weekBackgroundColour,false,false);//fontSize,Text Rotation, BackgroundColour
      
      //Format the day row 5 the same as the week row 3
      dayRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange).offset(2,0);
      formatRange(dayRange,8,0,weekBackgroundColour,false,false);//fontSize,Text Rotation, BackgroundColour
      
     //Format week divider to create a plan on a page POAP for a .ppt or google slide presentation
     let formatWeekRightBoarder = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(6,columnIndex,45);
     formatWeekRightBoarder.setBorder(null, true, null, null, null, null, '#434343', SpreadsheetApp.BorderStyle.SOLID);
    //Store the start values for the next range

      weekStartAddress = currentCellAddress;
      oldWeekNumber = currentWeekNumber;
      toggleWeekColour = !toggleWeekColour; //toggle the background colour for the next range
    } //end of  if (currentWeekNumber != oldWeekNumber)

      /************************************************************************************* 
                Weekend Change Trigger (i.e. Saturday or Sunday)
      **************************************************************************************/
      let planOnAPage = 45;// this is the height of the overall POAP roadmap canvas which fits on a MS- ppt or google slide
      
      //if the day starts with "S" i.e. a Saturday or Sunday then shade those columns 
      if (currentDayName.substring(0,1) === "S") {
        let columnHighLightRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(6,columnIndex,planOnAPage);
        columnHighLightRange.setBackground('#cfe2f3').setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
      }

    
    /************************************************************************************
     *         If the last item in the loop then tidy up the last cells in the ranges
    *************************************************************************************/
    if (j===numberOfDaysInRange-1){
      console.log("tidy weeks");
      currentYearRange=(yearStartAddress +":" + currentCellAddress );
      currentMonthRange=(monthStartAddress +":" + currentCellAddress );
      currentWeekRange=(weekStartAddress +":" + currentCellAddress );

      console.log(currentWeekRange);

      //Format the Week row cell background
      // set cell background for the current end of the range to colour red

      yearBackgroundColour= toggleBackgroundColour(toggleYearColour)
      monthBackgroundColour= toggleBackgroundColour(toggleMonthColour)
      weekBackgroundColour= toggleBackgroundColour(toggleWeekColour)
        

      //Format the last week range
      //format the week row background and text 
      
      //Format the YEAR row 1  
      yearRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentYearRange).offset(-2,0);
      formatRange(yearRange,14,0,yearBackgroundColour,true,true);//fontSize,Text Rotation, BackgroundColour, merge, bold
      
      //Format the MONTH row 2 
      monthRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentMonthRange).offset(-1,0);
      formatRange(monthRange,12,0,monthBackgroundColour,true,true);//fontSize,Text Rotation, BackgroundColour, merge bold

      //format the week row 3 
      weekRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange);
      formatRange(weekRange,9,0, weekBackgroundColour,true,false);//fontSize,Text Rotation, BackgroundColour, merge, bold
    
      //format the date range row 4 with  text orientation as 90 degrees
      dateRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange).offset(1,0);
      formatRange(dateRange,8,90, weekBackgroundColour,false,false);//fontSize,Text Rotation, BackgroundColour, merge, bold
      
      //Format the DAY row 5 
      // Try to offset using something like this dayRange=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1,canvasOffset_j,dimensions[0],dimensions[1]);
      dayRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange).offset(2,0);
      formatRange(dayRange,8,0, weekBackgroundColour,false,false);//fontSize,Text Rotation, BackgroundColour, merge bold

    }
      
    //currentDay, CurrentMonth, etc refers to the current month that's being calcuated in the loop

        dateArray.push([
                        currentYear,                      //Year
                        currentMonth,                     //Month
                        "WK_"+ currentWeekNumber[1],      //Week - select the 2nd value as the getISOWeekNumber() function returns year+week
                        currentDate,                      //Date
                        currentDayName.substring(0,1),    //Day of Week
                        rowIndex,                         //Row Index (i)
                        columnIndex,                      //Column Index (j)
                        currentCellAddress,               //current Cell Address in "A1" format
                        currentWeekRange,                 //Week Range trigger in "A1:B1" format
                        ])
        currentDay++ //next day
        console.log(currentDay) //next day
  } //End of for loop
 

  
  /*---------------------------------------------------------------------
      Transposing dateArray[] by 90 Degress into a new -> tranposedArray[]
  ----------------------------------------------------------------------*/
  //Transpose the data array so the dates go along the columns rather than along the rows.
  let tranposedArray = dateArray.map((_, colIndex) => dateArray.map(row => row[colIndex]));


  

  /*---------------------------------------------------------------------
    Paint the  tranposedArray[] & sliced to the canvas
    Slice the 2D transposedArray to only show the 1st 5 rows of the array from 0->4
    i.e. the Year, Month, Week, Date and Day details
  
    https://stackoverflow.com/questions/51383031/slice-section-of-two-dimensional-array-in-javascript
  

  Obtain the Dimensions of the tranposedArray[] this is required for the slicing dimensions[0]
  This dynamically works out the size of the tranposed array and sets the range size accordingly.
  ---------------------------------------------------------------------- */

  let dimensions = [tranposedArray.length,tranposedArray.reduce((x, y) => Math.max(x, y.length), 0)];
  console.log("dimensions of the tranposed array")
  let rowStart = 0;
  let rowEnd = 4; //the first 4 rows (zero-indexed so 5 rows) are visible. change to 11 to show the complete array for debugging
  let columnStart = 0
  let columnEnd = dimensions[0]
  let slicedArray = tranposedArray.slice(rowStart, rowEnd + 1).map(i => i.slice(columnStart, columnEnd + 1))
  
  /* 
  This dynamically works out the size of the array to determine the range before pasting. 
  This sets the range size accordingly to the size of the sliced array.
  */
  dimensions = [slicedArray.length,slicedArray.reduce((x, y) => Math.max(x, y.length), 0)];
  console.log("dimensions of the sliced array")
  console.log(dimensions)

  let paintRange=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1,canvasOffset_j,dimensions[0],dimensions[1]);
  paintRange.setHorizontalAlignment("center")
  paintRange.setValues(slicedArray);

  /*---------------------------------------------------------------------
           Get the used Ranges as a basis for mass formating of text etc
  ---------------------------------------------------------------------- */
  //Format the ranges
  //Get the used range 
  let lastColumn = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange).getLastColumn()
  let lastRow = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange).getLastRow()
  console.log("Row/Column Size = " + lastRow,lastColumn)
  setDates();
  canvasOffset(1,10);
};


function addDaysCopy(date, days) {
  const dateCopy = new Date(date);
  dateCopy.setDate(date.getDate() + days);
  return dateCopy;
}


function helloWorld() {
  Browser.msgBox("Hello World!");
}

function formatRange(range,fontSize,textRotation,rangeBackgroundColour,merge,bold){
//range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentWeekRange);
  range
    .setBackground(rangeBackgroundColour)
    .setFontSize(fontSize)
    .setFontColor("white")
    .setVerticalAlignment("center")
    .setHorizontalAlignment("center")
    .setTextRotation(textRotation)
    .setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
    
    //merge the range if the flag is set.
    if (merge){
      range.mergeAcross();
    }
    
    //bold the range if the flag is set.
    if (bold) {
    range.setFontWeight(bold);
  }
    SpreadsheetApp.getActive().getActiveSheet().setColumnWidths(1, 99, 15);
}

//function getWeek(date) {
//  return Number(Utilities.formatDate(new Date(date), "Europe/Kiev", "u")) === 7 ? 
//    Number(Utilities.formatDate(new Date(date), "Europe/Kiev", "w")) - 1 : 
//    Number(Utilities.formatDate(new Date(date), "Europe/Kiev", "w"));
//}


function toggleBackgroundColour(toggleColour){

 if (toggleColour) {
  rangeBackgroundColour = RGB1;
  // console.log("in the true part and toggleMonth = " + toggleColour + " = " + rangeBackgroundColour)
  //toggle = false;
}
else {
  rangeBackgroundColour = RGB2;
}
return rangeBackgroundColour;

}


function setDates() {
  //This functiona automatically identifies the timezone of the user
  var timezone = Session.getScriptTimeZone();
  Logger.log("YourTimeZone= " + timezone);
  
  var tempStartDate = new Date();
  tempStartDate.setDate(tempStartDate.getDate() - 7); //subtract 7 to get date from 1 week ago
  var startDate = Utilities.formatDate(tempStartDate, timezone, 'dd.MM.yyyy');
  Logger.log(startDate);

  var tempEndDate = new Date();
  var endDate = Utilities.formatDate(tempEndDate, timezone, 'dd.MM.yyyy');
  Logger.log(endDate);

  var dateRange = startDate + ' - ' + endDate;
  Logger.log(dateRange);
}



/* For a given date, get the ISO week number
 *
 * Based on information at:
 *
 *    THIS PAGE (DOMAIN EVEN) DOESN'T EXIST ANYMORE UNFORTUNATELY
 *    http://www.merlyn.demon.co.uk/weekcalc.htm#WNR
 *  UTC Date Demos Try are here - >  
 *  https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date/setMonth
 *
 * Algorithm is to find nearest thursday, it's year
 * is the year of the week number. Then get weeks
 * between that date and the first day of that year.
 *
 * Note that dates in one year can be weeks of previous
 * or next year, overlap is up to 3 days.
 *
 * e.g. 2014/12/29 is Monday in week  1 of 2015
 *      2012/1/1   is Sunday in week 52 of 2011
 */
function getISOWeekNumber(d) {
    // Copy date so don't modify original
    var d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    console.log("Target date in week function = " + d)
    // Set to nearest Thursday: current date + 4 - current day number
    // Make Sunday's day number 7
    let utcDate= d.getUTCDate();
    console.log("date  = " + utcDate)
    let utcDate2 = d.getUTCDay()||7
    console.log("Target Day Number (where 1=Monday) = " + utcDate2)

    let nearestThursday = d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
    console.log("year of nearest Thursday = " + d.toUTCString())
    // Get first day of year
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
    // Calculate full weeks to nearest Thursday
    var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
    // Return array of year and week number
    return [d.getUTCFullYear(), weekNo];
}

function canvasOffset(i,j) {

  //Get the roadmap sheet
  var roadMapSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
  //Determine the used range of the painted RoadMap
  let lastColumn= roadMapSheet.getLastColumn();
  let lastRow= roadMapSheet.getLastRow();
  console.log ("Row/Column = " + lastRow,lastColumn);
 

 //Move the canvas - currently fixed to 1,10  but could be variables for i,j
  roadMapSheet.getRange(1,1,lastRow+50,lastColumn).moveTo(roadMapSheet.getRange(i,j));

}

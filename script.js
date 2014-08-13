

function DoUpdateMetal()
{
  try
  {
    TimetableUpdate("Метал", 0) 
  }
  catch(err)
  {
    // Get the email address of the active user - that's you.
    var email = Session.getActiveUser().getEmail();  
    GmailApp.sendEmail(email, "UpdScript Falls", "There is fucking runtime error in script of updating Lemooor Studio Timetable. Check your tables!\n"+err);
  }
};





function DoUpdateRock()
{
  try
  {
    TimetableUpdate("Рок", 1) 
  }
  catch(err)
  {
    // Get the email address of the active user - that's you.
    var email = Session.getActiveUser().getEmail();  
    GmailApp.sendEmail(email, "UpdScript Falls", "There is fucking runtime error in script of updating Lemooor Studio Timetable. Check your tables!\n"+err);
  }
};





function DoUpdateJazz()
{
  try
  {
    TimetableUpdate("Джаз", 2) 
  }
  catch(err)
  {
    // Get the email address of the active user - that's you.
    var email = Session.getActiveUser().getEmail();  
    GmailApp.sendEmail(email, "UpdScript Falls", "There is fucking runtime error in script of updating Lemooor Studio Timetable. Check your tables!\n"+err);
  }
};



function DoUpdateChanson()
{
  try
  {
    TimetableUpdate("Шансон", 3) 
  }
  catch(err)
  {
    // Get the email address of the active user - that's you.
    var email = Session.getActiveUser().getEmail();  
    GmailApp.sendEmail(email, "UpdScript Falls", "There is fucking runtime error in script of updating Lemooor Studio Timetable. Check your tables!\n"+err);
  }
};




function TimetableUpdate(room, num) 
{ 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  TimetableUpdateProc(room, ss.getSheets()[num].getDataRange().getValues());
};

function TimetableUpdateProc(room, rows)
{
  var year = rows[0][0];
  var month = rows[1][0];
  var intervals = ["09:00:00", "12:00:00", "15:00:00", "18:00:00", "21:00:00", "24:00:00"];
  for(var i=0; i<(intervals.length-1); i++)
  {
    RowUpdProc(room, rows[i+2], intervals[i], intervals[i+1], year, month);
  }
};

function RowUpdProc(room, row, from, to, year, month)
{
 for(var i=1; i<row.length; i++)
 {
   var timestamp_begin = month + " " + i + ", " + year + " " + from
   var timestamp_end = month + " " + i + ", " + year + " " + to
   var calendar = CalendarApp.getCalendarsByName(room)[0];
   var rep = calendar.getEvents(new Date(timestamp_begin),new Date(timestamp_end))[0]
   if( CheckEmpty(row[i])  )
   {
     DeleteRep( rep )
   }
   else
   {
     CreateRep( rep, calendar, timestamp_begin, timestamp_end )
   }
 }
}

function CheckEmpty(str)
{
 return (str.search(/[А-яЁёA-z01-9]/) === -1)
}

function CreateRep( rep, calendar, timestamp_begin, timestamp_end )
{
  if(rep == undefined)
  {
    calendar.createEvent('Репетиция', new Date(timestamp_begin),new Date(timestamp_end))
  }
}

function DeleteRep( rep )
{
  if(rep != undefined)
  {
    rep.deleteEvent()
  }
}
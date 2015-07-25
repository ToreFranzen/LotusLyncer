using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Domino;

namespace LotusLyncer
{
    class NotesCalendar
    {
        public CalendarEvent GetCurrentMeeting(string password)
        {
            NotesSession notesSession = new NotesSession();
            NotesDatabase notesDatabase;
            
            try
            {
                notesSession.Initialize(password);               
                string MailServer = notesSession.GetEnvironmentString("MailServer", true);
                string MailFile = notesSession.GetEnvironmentString("MailFile", true);
                notesDatabase = notesSession.GetDatabase(MailServer, MailFile, false);

                NotesDateTime minStartDate = notesSession.CreateDateTime("Today");
                NotesDateTime maxEndDate = notesSession.CreateDateTime("Tomorrow");
                // Query Lotus Notes to get calendar entries in our date range.
                // To understand this SELECT, go to http://publib.boulder.ibm.com/infocenter/domhelp/v8r0/index.jsp
                // and search for the various keywords. Here is an overview:
                //   !@IsAvailable($Conflict) will exclude entries that conflicts with another.
                //     LN doesn't show conflict entries and we should ignore them.
                //   @IsAvailable(CalendarDateTime) is true if the LN document is a calendar entry
                //   @Explode splits a string based on the delimiters ",; "
                //   The operator *= is a permuted equal operator. It compares all entries on
                //     the left side to all entries on the right side. If there is at least one
                //     match, then true is returned.
                String calendarQuery = "SELECT (!@IsAvailable($Conflict) & @IsAvailable(CalendarDateTime) & (@Explode(CalendarDateTime) *= @Explode(@TextToTime(\""
                    + minStartDate.LocalTime
                    + "-" + maxEndDate.LocalTime + "\"))))";

                NotesDocumentCollection appointments = notesDatabase.Search(calendarQuery, null, 25);
                //appointments.Count
                NotesDocument Current = appointments.GetFirstDocument();
                DateTime rightNow = DateTime.Now;
                CalendarEvent closestMeeting = null;
                TimeSpan ts = new TimeSpan(2, 0, 0, 0); //set for 2 days as longest time span

                while (Current != null)
                {
                    int repeatCount = ((object[])Current.GetItemValue("StartDateTime")).Length;
                    for (int i = 0; i < repeatCount; i++)
                    {
                        var calendarEvent = GetCallendarEvent(Current, i);                        
                        if (calendarEvent.Starts <= rightNow && calendarEvent.Ends > rightNow)
                        {
                            //ding ding ding, winner                   
                            return calendarEvent;
                        }
                        TimeSpan timeToMeeting = calendarEvent.Starts.Subtract(rightNow);
                        if (timeToMeeting < ts && calendarEvent.Starts > rightNow)
                        {
                            ts = timeToMeeting;
                            closestMeeting = calendarEvent;
                        }
                    }

                    Current = appointments.GetNextDocument(Current);
                }
                return closestMeeting;

            }
            catch (Exception)
            {
                //TODO: Catch bad stuff here and provide a message
                throw;
            }
            
        }
        private CalendarEvent GetCallendarEvent(NotesDocument Current, int RepeatIndex)
        {
            CalendarEvent e = new CalendarEvent();
            try
            {                                

                e.Starts = (DateTime)((object[])Current.GetItemValue("StartDateTime"))[RepeatIndex];
                e.Ends = (DateTime)((object[])Current.GetItemValue("EndDateTime"))[RepeatIndex];
                e.Title = (string)((object[])Current.GetItemValue("Subject"))[0];
                e.Location = (string)((object[])Current.GetItemValue("Room"))[0] + " " +
                                (string)((object[])Current.GetItemValue("Location"))[0];
            }
            catch (Exception ex)
            {
               throw ex;
            }
            return e;
        }

        
    }
}

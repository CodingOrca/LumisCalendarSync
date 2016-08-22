using System.Collections.Generic;

using Microsoft.Office365.OutlookServices;


namespace LumisCalendarSync.Model
{
    public class OutlookCalendar
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
    }

    public class CalendarJsonData
    {
        public IList<Calendar> value { get; set; }
    }

    public class EventsJsonData
    {
        public IList<Event> value { get; set; }
    }
}

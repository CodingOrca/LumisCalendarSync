using System;
using Microsoft.Office365.OutlookServices;


namespace LumisCalendarSync.ViewModels
{
    public class EventModel: BindableBase
    {
        private readonly IEvent myEvent;

        public EventModel(IEvent e)
        {
            myEvent = e;
        }

        public IEvent Event
        {
            get { return myEvent; }
        }

        public string Subject
        {
            get { return Event.Subject; }
        }

        public string Location
        {
            get { return Event.Location.DisplayName; }
        }

        public string ShowAs
        {
            get { return Event.ShowAs.ToString(); }
        }

        public string IsRecurring
        {
            get { return Event.Type == EventType.SeriesMaster ? "yes" : "no"; }
        }

        public string Recurrence
        {
            get { return Event.Type == EventType.SeriesMaster ? Event.Recurrence.Pattern.Type.ToString() : "n.a."; }
        }

        public string Start
        {
            get
            {
                var dt = DateTime.Parse(Event.Start.DateTime);
                var timeZoneInfo = TimeZoneInfo.Utc;
                if (!string.IsNullOrWhiteSpace(Event.Start.TimeZone))
                {
                    timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(Event.Start.TimeZone);
                }
                dt = TimeZoneInfo.ConvertTime(dt, timeZoneInfo, TimeZoneInfo.Local);
                return dt.ToString("g");
            }
        }

        public string End
        {
            get
            {
                var dt = DateTime.Parse(Event.End.DateTime);
                var timeZoneInfo = TimeZoneInfo.Utc;
                if (!string.IsNullOrWhiteSpace(Event.End.TimeZone))
                {
                    timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(Event.End.TimeZone);
                }
                dt = TimeZoneInfo.ConvertTime(dt, timeZoneInfo, TimeZoneInfo.Local);
                return dt.ToString("g");
            }
        }

        private bool myIsSynchronized;
        public bool IsSynchronized
        {
            get { return myIsSynchronized; }
            set { Set(ref myIsSynchronized, value, "IsSynchronized"); }
        }

        public int? Reminder
        {
            get { return Event.ReminderMinutesBeforeStart; }
        }
    }
}

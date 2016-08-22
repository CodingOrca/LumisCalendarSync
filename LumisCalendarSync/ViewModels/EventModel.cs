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
            get { return myEvent.Subject; }
        }

        public string Location
        {
            get { return myEvent.Location.DisplayName; }
        }

        public string IsRecurring
        {
            get { return myEvent.Type == EventType.SeriesMaster ? "yes" : "no"; }
        }

        public string Recurrence
        {
            get { return myEvent.Type == EventType.SeriesMaster ? myEvent.Recurrence.Pattern.Type.ToString() : "n.a."; }
        }

        public string Start
        {
            get
            {
                return myEvent.Start.DateTime.Substring(0, 16) + " UTC";
            }
        }

        public string End
        {
            get
            {
                return myEvent.End.DateTime.Substring(0, 16) + " UTC";
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
            get { return myEvent.ReminderMinutesBeforeStart; }
        }
    }
}

using System;

namespace LumisCalendarSync.ViewModels
{
    public delegate void NotificationEventHandler(object sender, NotificationEventArgs e);

    public class NotificationEventArgs: EventArgs
    {
        public string Text { get; private set; }

        public NotificationEventArgs(string text)
        {
            Text = text;
        }
    }
}

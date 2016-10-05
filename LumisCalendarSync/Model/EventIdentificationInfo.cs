using System;
using System.Collections.Generic;


namespace LumisCalendarSync.Model
{
    public class EventIdentificationInfo
    {
        public string Id { get; set; }
        public String LastSyncTimeStamp { get; set; }
        public Dictionary<string, EventIdentificationInfo> ExceptionIds { get; set; }

        public EventIdentificationInfo()
        {
            ExceptionIds = new Dictionary<string, EventIdentificationInfo>();
        }
    }
}

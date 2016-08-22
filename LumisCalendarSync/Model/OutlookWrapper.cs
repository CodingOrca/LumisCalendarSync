using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

using Outlook = Microsoft.Office.Interop.Outlook; 

namespace LumisCalendarSync.Model
{
    internal sealed class OutlookWrapper : IDisposable
    {
        public OutlookWrapper()
        {
            if (Process.GetProcessesByName("OUTLOOK").Any())
            {
                try
                {
                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                    myOutlook = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                }
                catch
                {
                    myOutlook = null;
                }
            }
        }

        public bool IsOutlookRunning
        {
            get { return myOutlook != null; }
        }

        public Outlook.Items GetAppointmentItems()
        {
            if (myOutlook == null) return null;

            var defaultCal = myOutlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
            if (defaultCal == null)
            {
                return null;
            }
            var srcAppointmentItems = defaultCal.Items;
            Marshal.ReleaseComObject(defaultCal);
            return srcAppointmentItems;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~OutlookWrapper()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
            }
            if (myOutlook == null) return;

            Marshal.ReleaseComObject(myOutlook);
            myOutlook = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private Outlook.Application myOutlook;

    }
}

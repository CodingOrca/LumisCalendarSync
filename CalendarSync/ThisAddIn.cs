using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;

namespace LumisCalendarSync
{
    public partial class ThisAddIn
    {
        private Office.CommandBar menuBar;
        private Office.CommandBarPopup newMenuBar;
        private Office.CommandBarButton syncButton;
        private Office.CommandBarButton cleanupButton;
        private Office.CommandBarButton configButton;
        private Office.CommandBarButton logButton;

        private Timer timer;

        StreamWriter messageLog = null;
        string messageLogFileName = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            timer = new Timer();
            timer.Tick += timer_Tick;
            timer.Interval = 60000;
            timer.Start();

            messageLogFileName = Path.Combine(Path.GetTempPath(), "LumisCalendarSync.txt");
            messageLog = new StreamWriter(messageLogFileName);
            messageLog.AutoFlush = true;
            RegisterEvents();
            AddMenuBar();
        }

        private Outlook.Reminders myReminders;
        private Microsoft.Office.Interop.Outlook.SyncObject syncObject;

        private void RegisterEvents()
        {
            myReminders = this.Application.Reminders;
            myReminders.ReminderFire += new Outlook.ReminderCollectionEvents_ReminderFireEventHandler(myReminders_ReminderFire);

            Microsoft.Office.Interop.Outlook.SyncObjects syncObjects = this.Application.Session.SyncObjects;
            if (syncObjects.Count <= 0)
            {
                MessageBox.Show("You have to set up a sync group in Outlook");
            }
            else
            {
                syncObject = syncObjects[1];
                syncObject.SyncStart += syncObject_SyncStart;
                syncObject.SyncEnd += syncObject_SyncEnd;
                syncObject.OnError += syncObject_OnError;
            }
        }

        void syncObject_SyncStart()
        {
            timer.Stop();
            // if this happens in the time when an appointment sync was set up but not yet executed, we cancel that appointment sync by changing the timer Command to none:
            // The use case is most probably when an automatic connector sync was done few secconds ago and now the user pressed F9.
            // The syncObject_endSync will set up another AppointmentSync, so no problem in cancelling this one.
            if (timerCommand == TimerCommand.AppointmentsSync)
            {
                timerCommand = TimerCommand.None;
            }
        }

        void syncObject_OnError(int Code, string Description)
        {
            messageLog.WriteLine("Outlook Connector Error Code 0x{0:X}: {1}", Code, Description);
        }

        private enum TimerCommand { None, ConnectorSync, AppointmentsSync};

        TimerCommand timerCommand = TimerCommand.None;

        void syncObject_SyncEnd()
        {
            messageLog.WriteLine("{0}: Connector finished syncing.", DateTime.Now);
            switch (timerCommand)
            {
                case TimerCommand.None:
                    messageLog.WriteLine("   Setting up appointment sync.");
                    timerCommand = TimerCommand.AppointmentsSync;
                    timer.Start();
                    break;
                case TimerCommand.ConnectorSync:
                    timerCommand = TimerCommand.None;
                    break;
            }

        }

        void timer_Tick(object sender, EventArgs e)
        {
            timer.Stop();
            timer.Interval = 10000;
            switch (timerCommand)
            {
                case TimerCommand.AppointmentsSync:
                    messageLog.WriteLine("{0}: Start syncing appointments", DateTime.Now);
                    SyncAppointments();
                    timerCommand = TimerCommand.ConnectorSync;
                    timer.Start();
                    break;
                case TimerCommand.ConnectorSync:
                    if (syncObject != null)
                    {
                        messageLog.WriteLine("{0}: Starting Connector sync", DateTime.Now);
                        messageLog.WriteLine();
                        syncObject.Start();
                    }
                    else
                    {
                        messageLog.WriteLine("{0}: No sync group in Outlook to be triggered", DateTime.Now);
                    }
                    break;
                case TimerCommand.None:
                    if (syncObject != null)
                    {
                        messageLog.WriteLine("{0}: Initiating first sync after outlook has started.", DateTime.Now); 
                        syncObject.Start();
                    }
                    break;
            }
        }

        void myReminders_ReminderFire(Outlook.Reminder ReminderObject)
        {
            // is it not an appointment? Continue then.
            Microsoft.Office.Interop.Outlook.AppointmentItem appointment = ReminderObject.Item;
            if (appointment == null) return;

            Outlook.AppointmentItem parent = null;
            string subject = "unknown";
            try
            {
                subject = appointment.Subject;
                bool isSyncItem = appointment.UserProperties["originalId"] != null;
                if (!isSyncItem)
                {
                    // might be an instance of a recurring appointment:
                    parent = appointment.Parent as Outlook.AppointmentItem;
                    if (parent != null)
                    {
                        isSyncItem = parent.UserProperties["originalId"] != null;
                        Marshal.ReleaseComObject(parent);
                    }
                }
                Marshal.ReleaseComObject(appointment);
                if (isSyncItem)
                {
                    ReminderObject.Dismiss();
                    messageLog.WriteLine("{0}: Automatically dismissed the reminder for appointment [{1}], as it is a synced item", DateTime.Now, subject);
                }
            }
            catch (Exception)
            {
                messageLog.WriteLine("{0}: WARNING: could not clear the reminder for your synced appointment [{1}]", DateTime.Now, subject);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            messageLog.Close();
            messageLog.Dispose();
        }

        private void AddMenuBar()
        {
            try
            {
                menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
                if (newMenuBar != null)
                {
                    newMenuBar.Caption = "Lumis Calendar Sync";

                    syncButton = (Office.CommandBarButton)newMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    syncButton.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    syncButton.Caption = "Sync Now";
                    syncButton.FaceId = 37;
                    syncButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(syncButton_Click);

                    cleanupButton = (Office.CommandBarButton)newMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 2, true);
                    cleanupButton.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    cleanupButton.Caption = "Clean up target calendar";
                    cleanupButton.FaceId = 358;
                    cleanupButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cleanupButton_Click);

                    configButton = (Office.CommandBarButton)newMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 3, true);
                    configButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                    configButton.Caption = "Settings";
                    configButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(configButton_Click);

                    logButton = (Office.CommandBarButton)newMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 4, true);
                    logButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                    logButton.Caption = "View Log file";
                    logButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(logButton_Click);

                    newMenuBar.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void SyncNow()
        {
            if (syncObject != null)
            {
                syncObject.Start();
            }
        }

        private void syncButton_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            if (timer.Enabled)
            {
                messageLog.WriteLine("{0}: Manually triggered sync not performed because an automatic sync is in progress.", DateTime.Now);
                return;
            }
            messageLog.WriteLine("{0}: Manually triggered sync", DateTime.Now);
            timerCommand = TimerCommand.AppointmentsSync;
            timer.Start();
        }

        private void cleanupButton_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            StringBuilder sb1 = new StringBuilder();
            sb1.AppendLine(String.Format("This will delete all synced appointments from [{0}] and will disable automatic sync to this calendar.", Properties.Settings.Default.DestinationCalendar));
            sb1.AppendLine("You can set up Automatic sync to the same calendar or another calendar afterwards. Your default calendar is not affected.");
            var response = MessageBox.Show(sb1.ToString(), "Lumis Calendar Sync", MessageBoxButtons.OKCancel);
            if (response == DialogResult.Cancel) return;

            if (timer.Enabled)
            {
                timer.Stop();
                timerCommand = TimerCommand.None;
            }
            messageLog.WriteLine("{0}: Cleaning up target calendar", DateTime.Now);
            int deletedItems = this.CleanupTargetCalendar();
            string targetCalendar = Properties.Settings.Default.DestinationCalendar;
            Properties.Settings.Default.DestinationCalendar = "None";
            Properties.Settings.Default.Save();
            timerCommand = TimerCommand.ConnectorSync;
            timer.Start();
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("{0} synced Appointments deleted from calendar [{1}].", deletedItems, targetCalendar));
            sb.AppendLine("Automatic sync is disabled. Use Lumis Calendar Sync add-in Settings Menu to set up sync to a new target calendar.");
            MessageBox.Show(sb.ToString(), "Lumis Calendar Sync");
        }

        private void configButton_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            var configForm = new ConfigurationForm(this);
            configForm.Show();
        }

        private void logButton_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            System.Diagnostics.Process.Start(messageLogFileName);
        }

        public List<string> GetPossibleDestinationCalendars()
        {
            List<string> result = new List<string>();
            result.Add("None");

            var allRootFolders = this.Application.Session.Folders;
            Outlook.Folder defaultCal = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
            string defaultCalendarPath = "";
            if (defaultCal != null )
            {
                defaultCalendarPath = defaultCal.FolderPath;
            }

            foreach (Outlook.Folder folder in allRootFolders)
            {
                foreach (Outlook.Folder subFolder in folder.Folders)
                {
                    if (subFolder != null && subFolder.DefaultItemType == Outlook.OlItemType.olAppointmentItem && defaultCalendarPath != subFolder.FolderPath )
                    {
                        result.Add(String.Format("{0} in {1}", subFolder.Name, folder.Name));
                    }
                }
            }
            return result;
        }

        private int CleanupTargetCalendar()
        {
            // set the dstAppointmetItems to the items of the destionation Calendar:
            Outlook.Items dstAppointmentItems = GetAllTargetAppointments();

            // get all appointments in the dstAppointmentItems which where synced before:
            Dictionary<string, Outlook.AppointmentItem> targetItems = GetAllSyncedAppointments(dstAppointmentItems);

            messageLog.WriteLine("Deleting all synced appointments in the target calendar to enforce a sync.");
            int deletedItems = DeleteAppointments(targetItems);

            messageLog.WriteLine("{0}: Cleanup done", DateTime.Now);
            if (targetItems.Count != deletedItems)
            {
                messageLog.WriteLine("{0} items could not be deleted!", targetItems.Count - deletedItems);
            }
            messageLog.WriteLine("{0} items deleted", deletedItems);
            messageLog.WriteLine();
            Marshal.ReleaseComObject(dstAppointmentItems);
            return deletedItems;
        }

        private void SyncAppointments()
        {
            if (LumisCalendarSync.Properties.Settings.Default.DestinationCalendar == "None" || String.IsNullOrWhiteSpace(LumisCalendarSync.Properties.Settings.Default.DestinationCalendar))
            {
                messageLog.WriteLine("No Destination Calendar is selected, sync is disabled.");
                messageLog.WriteLine();
                return;
            }
            int oldAppointments = 0;
            int unchangedAppointments = 0;
            int successfullyUpdated = 0;
            int errorUpdated = 0;
            int deletedUpdates = 0;

            int srcAppointmentReferences = 0;
            int dstAppointmentReferences = 0;

            string currentSubject = "no appointment processed yet";
            string operationChain = "";

            Outlook.Items dstAppointmentItems = null;
            Outlook.Items srcAppointmentItems = null;

            try
            {
                #region set the srcAppointmentItems to the Items of default calendar:
                {
                    Outlook.Folder defaultCal = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
                    srcAppointmentItems = defaultCal.Items;
                    Marshal.ReleaseComObject(defaultCal);
                    defaultCal = null;
                }
                #endregion

                // set the dstAppointmetItems to the items of the destionation Calendar:
                dstAppointmentItems = GetAllTargetAppointments();

                operationChain = "listing all appointments in target calendar";

                // get all appointments in the dstAppointmentItems which where synced before:
                Dictionary<string, Outlook.AppointmentItem> targetItems = GetAllSyncedAppointments(dstAppointmentItems);

                dstAppointmentReferences += targetItems.Count;

                foreach (var item in srcAppointmentItems)
                {
                    Outlook.AppointmentItem srcAppointment = null;
                    try
                    {
                        srcAppointment = item as Outlook.AppointmentItem;
                    }
                    catch
                    {
                        srcAppointment = null;
                    }

                    if (srcAppointment == null) continue;

                    srcAppointmentReferences++;

                    try
                    {
                        currentSubject = srcAppointment.Subject;
                        operationChain = "Checking source end date; ";
                        // skip non-recurring appintments which ended more than 90 days ago
                        if (!srcAppointment.IsRecurring && srcAppointment.End < DateTime.Now - TimeSpan.FromDays(90))
                        {
                            oldAppointments++;
                            continue;
                        }
                        // skip recurring appintments with last occurance more than 90 days ago
                        if (srcAppointment.IsRecurring)
                        {
                            var srcPattern = srcAppointment.GetRecurrencePattern();
                            if (!srcPattern.NoEndDate && srcPattern.PatternEndDate < DateTime.Now - TimeSpan.FromDays(90))
                            {
                                oldAppointments++;
                                continue;
                            }
                        }

                        Outlook.AppointmentItem dstAppointment = null;
                        operationChain += "Checking if target appointment alreay exists; ";
                        if (targetItems.ContainsKey(srcAppointment.GlobalAppointmentID))
                        {
                            dstAppointment = targetItems[srcAppointment.GlobalAppointmentID];
                            targetItems.Remove(srcAppointment.GlobalAppointmentID);
                            if (dstAppointment.UserProperties["originalLastUpdate"] != null)
                            {
                                string lastSyncTime = dstAppointment.UserProperties["originalLastUpdate"].Value as string;
                                // skip appointments which did not changed since last sync
                                if (lastSyncTime == srcAppointment.LastModificationTime.ToString())
                                {
                                    unchangedAppointments++;

                                    Marshal.ReleaseComObject(dstAppointment);
                                    dstAppointmentReferences--;

                                    continue;
                                }
                            }
                            // appointments for which reccurance or AllDayEvent changed, we delete the target appointment (and create a new one later down).
                            if (dstAppointment.IsRecurring != srcAppointment.IsRecurring || srcAppointment.AllDayEvent != dstAppointment.AllDayEvent)
                            {
                                dstAppointment.Delete();
                                dstAppointment = null;
                                dstAppointmentReferences--;
                            }
                        }

                        // this will indicate if we create a new appointment in the target folder
                        bool dstAppointmentIsNew = false;
                        bool retry = true;

                        while (retry)
                        {
                            retry = false;
                            try
                            {
                                messageLog.Write("  Updating [{0}] ", srcAppointment.Subject);

                                if (dstAppointment == null)
                                {
                                    dstAppointmentIsNew = true;
                                    operationChain += "Creating a new target appointment; ";
                                    dstAppointment = dstAppointmentItems.Add(Outlook.OlItemType.olAppointmentItem);
                                    dstAppointmentReferences++;

                                    dstAppointment.UserProperties.Add("originalId", Outlook.OlUserPropertyType.olText, false);
                                    dstAppointment.UserProperties["originalId"].Value = srcAppointment.GlobalAppointmentID;
                                    dstAppointment.UserProperties.Add("originalLastUpdate", Outlook.OlUserPropertyType.olText, false);
                                    dstAppointment.UserProperties["originalLastUpdate"].Value = srcAppointment.LastModificationTime.ToString();
                                    if (srcAppointment.AllDayEvent)
                                    {
                                        operationChain += "AllDayEvent; ";
                                        dstAppointment.AllDayEvent = true;
                                    }
                                }

                                operationChain += "Updating Subject; ";
                                dstAppointment.Subject = srcAppointment.Subject;
                                operationChain += "Updating Location; ";
                                dstAppointment.Location = srcAppointment.Location;
                                operationChain += "Updating BusyStatus; ";
                                dstAppointment.BusyStatus = srcAppointment.BusyStatus;

                                if (!srcAppointment.IsRecurring)
                                {
                                    messageLog.WriteLine("on {0}", srcAppointment.Start);

                                    operationChain += "Not Recurring; ";

                                    operationChain += "Updating Start; ";
                                    dstAppointment.Start = srcAppointment.Start;

                                    operationChain += "Updating Duration; ";
                                    dstAppointment.Duration = srcAppointment.Duration;

                                    operationChain += "Updating originalLastUpdate; ";

                                    dstAppointment.UserProperties["originalLastUpdate"].Value = srcAppointment.LastModificationTime.ToString();

                                    operationChain += "Saving; ";
                                    dstAppointment.Save();
                                    // messageLog.WriteLine("  Non-reccuring appoitment [{0}] was succesfully synced", appointment.Subject);
                                    successfullyUpdated++;
                                }
                                else // IsRecurring
                                {
                                    operationChain += "Recurring; ";
                                    var srcPattern = srcAppointment.GetRecurrencePattern();
                                    var dstPattern = dstAppointment.GetRecurrencePattern();

                                    operationChain += "Updating RecurreneType; ";
                                    dstPattern.RecurrenceType = srcPattern.RecurrenceType;

                                    if (!srcAppointment.AllDayEvent)
                                    {
                                        operationChain += "Updating StartTime; ";
                                        dstPattern.StartTime = srcPattern.StartTime;
                                    }

                                    operationChain += "Updating Duration; ";
                                    dstPattern.Duration = srcPattern.Duration;

                                    messageLog.WriteLine("recurring {0} at {1} ", dstPattern.RecurrenceType, dstPattern.StartTime);

                                    switch (dstPattern.RecurrenceType)
                                    {
                                        case Outlook.OlRecurrenceType.olRecursDaily:
                                            break;
                                        case Outlook.OlRecurrenceType.olRecursWeekly:
                                            dstPattern.DayOfWeekMask = srcPattern.DayOfWeekMask;
                                            break;
                                        case Outlook.OlRecurrenceType.olRecursMonthly:
                                            dstPattern.DayOfMonth = srcPattern.DayOfMonth;
                                            break;
                                        case Outlook.OlRecurrenceType.olRecursMonthNth:
                                            dstPattern.DayOfWeekMask = srcPattern.DayOfWeekMask;
                                            dstPattern.Instance = srcPattern.Instance;
                                            break;
                                        case Outlook.OlRecurrenceType.olRecursYearly:
                                            dstPattern.DayOfMonth = srcPattern.DayOfMonth;
                                            dstPattern.MonthOfYear = srcPattern.MonthOfYear;
                                            break;
                                        case Outlook.OlRecurrenceType.olRecursYearNth:
                                            dstPattern.DayOfWeekMask = srcPattern.DayOfWeekMask;
                                            dstPattern.Instance = srcPattern.Instance;
                                            break;

                                    }

                                    dstPattern.PatternStartDate = srcPattern.PatternStartDate;
                                    if (srcPattern.Interval > 0)
                                    {
                                        dstPattern.Interval = srcPattern.Interval;
                                    }
                                    dstPattern.NoEndDate = srcPattern.NoEndDate;
                                    if (!srcPattern.NoEndDate)
                                    {
                                        dstPattern.Occurrences = srcPattern.Occurrences;
                                        dstPattern.PatternEndDate = srcPattern.PatternEndDate;
                                    }

                                    operationChain += "Saving; ";
                                    dstAppointment.Save();

                                    Outlook.Exceptions srcExceptions = srcPattern.Exceptions;

                                    try
                                    {

                                        if (srcExceptions != null && srcExceptions.Count > 0)
                                        {
                                            messageLog.WriteLine("    Syncing {0} exceptions for this recurring appointment:", srcExceptions.Count);
                                        }
                                        else
                                        {
                                            messageLog.WriteLine("    This recurring appointment has no exceptions.");
                                        }


                                        foreach (Outlook.Exception srcException in srcExceptions)
                                        {
                                            DateTime originalDate = new DateTime(srcException.OriginalDate.Year, srcException.OriginalDate.Month, srcException.OriginalDate.Day,
                                                srcPattern.StartTime.Hour, srcPattern.StartTime.Minute, srcPattern.StartTime.Second);

                                            Outlook.AppointmentItem dstExceptionItem = null;
                                            Outlook.AppointmentItem srcExceptionItem = null;

                                            try
                                            {
                                                dstExceptionItem = dstPattern.GetOccurrence(originalDate);
                                                dstAppointmentReferences++;
                                                if (srcException.Deleted)
                                                {
                                                    messageLog.WriteLine("    On {0}: occurence deleted.", originalDate.ToShortDateString());
                                                    dstExceptionItem.Delete();
                                                    dstExceptionItem = null;
                                                    dstAppointmentReferences--;
                                                    continue;
                                                }

                                                srcExceptionItem = srcException.AppointmentItem;
                                                srcAppointmentReferences++;

                                                messageLog.WriteLine("    On {0}: shifting occurence to {1}.", originalDate.ToShortDateString(), srcExceptionItem.Start);

                                                dstExceptionItem.Subject = srcExceptionItem.Subject;
                                                dstExceptionItem.Location = srcExceptionItem.Location;
                                                dstExceptionItem.Start = srcExceptionItem.Start;
                                                dstExceptionItem.End = srcExceptionItem.End;

                                                dstExceptionItem.Save();
                                            }

                                            finally
                                            {
                                                if( dstExceptionItem != null )
                                                {
                                                    Marshal.ReleaseComObject(dstExceptionItem);
                                                    dstAppointmentReferences--;
                                                }
                                                if (srcExceptionItem != null)
                                                {
                                                    Marshal.ReleaseComObject(srcExceptionItem);
                                                    srcAppointmentReferences--;
                                                }
                                            }
                                        }
                                        dstAppointment.UserProperties["originalLastUpdate"].Value = srcAppointment.LastModificationTime.ToString();
                                        dstAppointment.Save();
                                        successfullyUpdated++;
                                    }
                                    finally
                                    {
                                        Marshal.ReleaseComObject(srcExceptions);
                                    }

                                } // the else of if(Recurring)
                            } 
                            catch (Exception ex)
                            {
                                if (dstAppointmentIsNew)
                                {
                                    messageLog.WriteLine("  ERROR: Could not create appointment [{0}] in target calendar.", srcAppointment.Subject);
                                    messageLog.WriteLine("  Chain of performed operations: {0}", operationChain);
                                    messageLog.WriteLine(ex.ToString());
                                    messageLog.WriteLine();
                                    errorUpdated++;
                                }
                                else
                                {
                                    messageLog.WriteLine("  WARNING: first atempt failed, retrying");
                                    dstAppointment.Delete();
                                    dstAppointment = null;
                                    dstAppointmentReferences--;
                                    retry = true;
                                }
                            }
                            finally
                            {
                                if (dstAppointment != null)
                                {
                                    Marshal.ReleaseComObject(dstAppointment);
                                    dstAppointment = null;
                                    dstAppointmentReferences--;
                                }
                            }
                        } // end of while(retry)
                    }
                    catch (Exception ex)
                    {
                        messageLog.WriteLine("ERROR syncing [{0}]. The message below might help us understanding what happened. Sorry.", currentSubject);
                        messageLog.WriteLine("Chain of performed operations: {0}", operationChain);
                        messageLog.WriteLine(ex.Message);
                        messageLog.WriteLine();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(srcAppointment);
                        srcAppointmentReferences--;
                    }
                }

                int deletedItems = DeleteAppointments(targetItems);

                deletedUpdates += deletedItems;
                errorUpdated += (targetItems.Count - deletedItems);
                dstAppointmentReferences -= deletedItems;

                messageLog.WriteLine("{0}: Sync done", DateTime.Now);
                messageLog.WriteLine("{0} appointments updated.", successfullyUpdated);
                messageLog.WriteLine("{0} appointments deleted.", deletedUpdates);
                messageLog.WriteLine("{0} appointments failed to be updated.", errorUpdated);
                messageLog.WriteLine("{0} appointments not synced because they ended more than 90 days ago.", oldAppointments);
                messageLog.WriteLine("{0} appointments not synced because they did not changed since their last sync.", unchangedAppointments);

            }
            catch (Exception ex)
            {
                messageLog.WriteLine("ERROR syncing. The message below might help us understanding what happened. Sorry.");
                messageLog.WriteLine("Appointment Subject: {0}", currentSubject);
                messageLog.WriteLine("Chain of performed operations: {0}", operationChain);
                messageLog.WriteLine(ex.Message);
            }
            finally
            {
                if (srcAppointmentItems != null)
                {
                    Marshal.ReleaseComObject(srcAppointmentItems);
                    srcAppointmentItems = null;
                }
                if (dstAppointmentItems != null)
                {
                    Marshal.ReleaseComObject(dstAppointmentItems);
                    dstAppointmentItems = null;
                }

                messageLog.WriteLine("Debug Information: srcReferences = {0}; dstReferences = {1}.", srcAppointmentReferences, dstAppointmentReferences);
                messageLog.WriteLine();
                messageLog.Flush();
            }
        }

        private Outlook.Items GetAllTargetAppointments()
        {
            Outlook.Folder targetCal = null;
            var allRootFolders = this.Application.Session.Folders;

            foreach (Outlook.Folder folder in allRootFolders)
            {
                foreach (Outlook.Folder subFolder in folder.Folders)
                {
                    if (subFolder != null && subFolder.DefaultItemType == Outlook.OlItemType.olAppointmentItem)
                    {
                        string calendarString = String.Format("{0} in {1}", subFolder.Name, folder.Name);
                        if (LumisCalendarSync.Properties.Settings.Default.DestinationCalendar == calendarString)
                        {
                            targetCal = subFolder;
                            messageLog.WriteLine("Your target calendar is [{0}]", calendarString);
                        }
                        else
                        {
                            Marshal.ReleaseComObject(subFolder);
                        }
                    }
                }
                Marshal.ReleaseComObject(folder);
            }

            if (targetCal == null)
            {
                throw new Exception(String.Format("Configured target calendar [{0}] could not be found", LumisCalendarSync.Properties.Settings.Default.DestinationCalendar));
            }

            Outlook.Items result = targetCal.Items;
            Marshal.ReleaseComObject(targetCal);
            targetCal = null;
            return result;
        }

        private int DeleteAppointments(Dictionary<string, Outlook.AppointmentItem> targetItems)
        {
            int deletedItems = 0;
            foreach (var item in targetItems.Values)
            {
                messageLog.WriteLine("  Deleting [{0}]", item.Subject);
                try
                {
                    item.Delete();
                    deletedItems++;
                }
                catch (Exception ex)
                {
                    messageLog.WriteLine("  ERROR: Could not delete appointment [{0}]: {1}", item.Subject, ex.Message);
                }
            }
            return deletedItems;
        }

        /// <summary>
        /// Returns a map with all appointments in the dstAppontmentItems which have been synced before.
        /// The key of an entry is the Appointment ID of the source appointment which was synced. The value it the target appointment itself.
        /// The synced appointments are identified by verifying if they have the User Property "originalID".
        /// </summary>
        /// <param name="dstAppointmentItems">The list of target appointments to be scanned</param>
        /// <returns></returns>
        private static Dictionary<string, Outlook.AppointmentItem> GetAllSyncedAppointments(Outlook.Items dstAppointmentItems)
        {
            Dictionary<string, Outlook.AppointmentItem> targetItems = new Dictionary<string, Outlook.AppointmentItem>();

            foreach (Outlook.AppointmentItem appointment in dstAppointmentItems)
            {
                if (appointment == null) continue;
                if (appointment.UserProperties["originalId"] != null)
                {
                    if (targetItems.ContainsKey(appointment.UserProperties["originalId"].Value as string))
                    {
                        // this should never happen. However, if this really happens we delete this appointment to avoid future errors
                        appointment.Delete();
                    }
                    else
                    {
                        targetItems.Add(appointment.UserProperties["originalId"].Value as string, appointment);
                    }
                }
                else
                {
                    Marshal.ReleaseComObject(appointment);
                }

            }
            return targetItems;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

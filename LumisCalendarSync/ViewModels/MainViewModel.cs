using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Threading;
using IWshRuntimeLibrary;
using LumisCalendarSync.Model;
using Microsoft.OData.ProxyExtensions;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office365.OutlookServices;

using DayOfWeek = Microsoft.Office365.OutlookServices.DayOfWeek;
using Exception = System.Exception;
using File = System.IO.File;
using RecurrencePattern = Microsoft.Office365.OutlookServices.RecurrencePattern;


namespace LumisCalendarSync.ViewModels
{
    public class MainViewModel: BindableBase
    {
        public MainViewModel()
        {
            OAuthHelper = new OAuthHelper();
            Calendars = new ObservableCollection<ICalendar>();
            Events = new ObservableCollection<EventModel>();
            LogEntries = new ObservableCollection<string>();
            
            myTimer = new DispatcherTimer();
            myTimer.Tick += Timer_Tick;
            myTimer.Interval = TimeSpan.FromMinutes(Properties.Settings.Default.AutoSyncInterval);

            if (InDesignMode())
            {
                LogFileName = "C:\\Temp\\SomeFileName.log";
                Events.Add(
                    new EventModel(new Event
                    {
                        Subject = "Designer Event Subject",
                        Location = new Location {DisplayName = "Wherever"},
                        Start = new DateTimeTimeZone {DateTime = DateTime.Now.ToString("O")},
                        End = new DateTimeTimeZone {DateTime = (DateTime.Now + TimeSpan.FromHours(1)).ToString("O")}
                    }));
                LogEntries.Add("Some Log entry");
                return;
            }

            var localStorageDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            myAppDataFolder = Path.Combine(localStorageDir, "LumisCalendarSync");
            if (!Directory.Exists(myAppDataFolder))
            {
                Directory.CreateDirectory(myAppDataFolder);
            }

            LogFileName = Path.Combine(myAppDataFolder, "LumisCalendarSync.txt");
            myMappingTable = new Dictionary<string, EventIdentificationInfo>();

            AutologinAsync();
        }

        public string LogFileName { get; private set; }

        public ObservableCollection<string> LogEntries { get; private set; }

        private void WriteMessageLog(string format, params object[] arguments)
        {
            // re-open the file if it exceeds 250 KB
            if (myMessageLog != null && myMessageLog.BaseStream.Length > 250 * 1024 * 1024)
            {
                myMessageLog.Close();
                myMessageLog = null;
            }
            if (myMessageLog == null)
            {
                try
                {
                    myMessageLog = new StreamWriter(LogFileName) { AutoFlush = true };
                }
                catch (Exception ex)
                {
                    Error = String.Format("Could not write to log file {0}, probably another instance of this app is running", LogFileName);
                    LogEntries.Add(ex.ToString());
                    return;
                }
            }

            string time = DateTime.Now.ToString("s");
            if (String.IsNullOrWhiteSpace(format))
            {
                LogEntries.Add(String.Format(""));
                myMessageLog.WriteLine();
            }
            else
            {
                LogEntries.Add(String.Format("{0}: {1}", time, String.Format(format, arguments)));
                myMessageLog.WriteLine("{0}: {1}", time, String.Format(format, arguments));
            }
        }

        private void SaveMappingTable()
        {
            var mappingFile = Path.Combine(myAppDataFolder, String.Format("Mapping-{0}.dat", SelectedCalendar.Id.GetHashCode()));
            try
            {
                var serializer = new JavaScriptSerializer();
                File.WriteAllText(mappingFile, serializer.Serialize(myMappingTable));
            }
            catch (Exception ex)
            {
                WriteMessageLog("Could not save Mapping file {0}: {1}", mappingFile, ex);
            }
        }

        private void LoadMappingTable()
        {
            var mappingFile = Path.Combine(myAppDataFolder, String.Format("Mapping-{0}.dat", SelectedCalendar.Id.GetHashCode()));
            if (!File.Exists(mappingFile))
            {
                myMappingTable.Clear();
                return;
            }
            try
            {
                var serializer = new JavaScriptSerializer();
                myMappingTable = serializer.Deserialize(File.ReadAllText(mappingFile), myMappingTable.GetType()) as Dictionary<string, EventIdentificationInfo>;
            }
            catch (Exception ex)
            {
                WriteMessageLog("Could not load Mapping file {0}: {1}", mappingFile, ex);
            }
        }

        private string myError;
        public string Error
        {
            get { return myError; }
            set
            {
                Set(ref myError, value, "Error"); 
                if (UserNotification != null && !String.IsNullOrWhiteSpace(value))
                {
                    UserNotification(this, new NotificationEventArgs(value));
                }
            }
        }

        public event NotificationEventHandler UserNotification;

        public ObservableCollection<ICalendar> Calendars { get; set; }

        public ObservableCollection<EventModel> Events { get; set; }

        private ICalendar mySelectedCalendar;
        public ICalendar SelectedCalendar
        {
            get { return mySelectedCalendar; }
            set
            {
                Set(ref mySelectedCalendar, value, "SelectedCalendar");
                if (SelectedCalendar != null)
                {
                    Properties.Settings.Default.RemoteCaleandarId = SelectedCalendar.Id;
                    Properties.Settings.Default.Save();
                    LoadMappingTable();
                    PopulateEventsAsync(SelectedCalendar.Id);
                }
                RaisePropertyChanged("CanAutosync");
                NotifyCommands();
            }
        }

        private EventModel mySelectedEvent;

        public EventModel SelectedEvent
        {
            get { return mySelectedEvent; }
            set
            {
                Set(ref mySelectedEvent, value, "SelectedEvent");
                NotifyCommands();
            }
        }

        public bool IsAutoSyncEnabled
        {
            get { return Properties.Settings.Default.IsAutoSyncEnabled; }
            set
            {
                Properties.Settings.Default.IsAutoSyncEnabled = value;
                Properties.Settings.Default.Save();
                RaisePropertyChanged("IsAutoSyncEnabled");
                RaisePropertyChanged("CanChangeCalendar");
                myTimer.IsEnabled = value;
            }
        }

        public int AutoSyncInterval
        {
            get { return Properties.Settings.Default.AutoSyncInterval; }
            set
            {
                if(value < 10 ) throw new Exception("Provide a value in minutes > 10");
                Properties.Settings.Default.AutoSyncInterval = value;
                Properties.Settings.Default.Save();
                myTimer.Interval = TimeSpan.FromMinutes(value);
                RaisePropertyChanged("AutoSyncInterval");
            }
        }

        public bool CanAutosync
        {
            get { return IsLoggedIn && SelectedCalendar != null; }
        }

        public bool CanChangeCalendar
        {
            get { return IsLoggedIn && IsIdle && !IsAutoSyncEnabled; }
        }

        private void NotifyCommands()
        {
            DeleteAllCommand.RaiseCanExecuteChanged();
            DeleteEventCommand.RaiseCanExecuteChanged();
            SynchronizeCommand.RaiseCanExecuteChanged();
        }

        private IUser myUser;

        public IUser User
        {
            get { return myUser; }
            set { Set(ref myUser, value, "User"); }
        }

        private int myRunningAsyncOperations;

        private int RunningAsyncOperations
        {
            get { return myRunningAsyncOperations; }
            set
            {
                //if (value < 0) value = 0;
                Set(ref myRunningAsyncOperations, value, "RunningAsyncOperations"); 
                RaisePropertyChanged("IsIdle");
                RaisePropertyChanged("CanChangeCalendar");
                NotifyCommands();
            }
        }

        public bool IsIdle
        {
            get { return RunningAsyncOperations == 0; }
        }

        private bool myIsLoggedIn;
        public bool IsLoggedIn
        {
            get { return myIsLoggedIn; }
            set
            {
                Set(ref myIsLoggedIn, value, "IsLoggedIn"); 
                RaisePropertyChanged("IsLoggedOut");
                RaisePropertyChanged("CanChangeCalendar");
                RaisePropertyChanged("CanAutosync");
                NotifyCommands();
            }
        }

        public bool IsLoggedOut
        {
            get { return !myIsLoggedIn; }
        }

        private DelegateCommand myDeleteAllCommand;
        public DelegateCommand DeleteAllCommand
        {
            get
            {
                if (myDeleteAllCommand == null)
                {
                    myDeleteAllCommand = DelegateCommand.FromAsyncHandler(
                        executeMethod: () => DeleteAllEventsAsync(),
                        canExecuteMethod: () => SelectedCalendar != null && IsIdle && IsLoggedIn);
                }
                return myDeleteAllCommand;
            }
        }

        private DelegateCommand myDeleteEventCommand;

        public DelegateCommand DeleteEventCommand
        {
            get
            {
                if (myDeleteEventCommand == null)
                {
                    myDeleteEventCommand = DelegateCommand.FromAsyncHandler(
                        executeMethod: () => DeleteEventAsync(SelectedEvent),
                        canExecuteMethod: () => SelectedEvent != null && IsIdle && IsLoggedIn);
                }
                return myDeleteEventCommand;
            }
        }

        private DelegateCommand mySynchronizeCommand;
        public DelegateCommand SynchronizeCommand
        {
            get
            {
                if (mySynchronizeCommand == null)
                {
                    mySynchronizeCommand = DelegateCommand.FromAsyncHandler(
                        executeMethod: SynchronizeAsync,
                        canExecuteMethod: () => SelectedCalendar != null && IsIdle && IsLoggedIn
                        );
                }
                return mySynchronizeCommand;
            }
        }

        public bool RunAtStartup
        {
            get { return IsStartupShortcutSaved(); }
            set
            {
                if(value) CreateStartupFolderShortcut();
                else DeleteStartupFolderShortcut();
                RaisePropertyChanged("RunAtStartup");
            }
        }

        async void Timer_Tick(object sender, EventArgs e)
        {
            await SynchronizeAsync();
        }

        async private Task SynchronizeAsync()
        {
            myTimer.Stop();
            Error = "";
            if (!IsLoggedIn)
            {
                Error = "User not logged in, cannot sync";
                return;
            }
            if (String.IsNullOrEmpty(Properties.Settings.Default.RemoteCaleandarId))
            {
                Error = "No remote calendar selected";
                return;
            }

            int oldAppointments = 0;
            int unchangedAppointments = 0;
            int successfullyUpdated = 0;
            int errorUpdated = 0;
            int deletedUpdates = 0;

            string currentSubject = "no appointment processed yet";
            string operationChain = "";

            Items srcAppointmentItems = null;
            RunningAsyncOperations++;
            try
            {
                using (var outlookWrapper = new OutlookWrapper())
                {
                    srcAppointmentItems = outlookWrapper.GetAppointmentItems();
                    if (srcAppointmentItems == null)
                    {
                        Error = "Outlook is not running, cannot sync";
                        return;
                    }

                    var remoteCalendarEvents = myOutlookServicesClient.Me.Calendars[Properties.Settings.Default.RemoteCaleandarId].Events;
                    Events.Clear();
                    LogEntries.Clear();
                    var dstAppointmentItems = await GetCalendarEventsAsync(Properties.Settings.Default.RemoteCaleandarId);

                    var targetItems = GetAllSyncedEvents(dstAppointmentItems);

                    foreach (var e in dstAppointmentItems)
                    {
                        if (targetItems.Values.Contains(e)) continue;
                        Events.Add(new EventModel(e) { IsSynchronized = false });
                    }

                    WriteMessageLog("Starting syncing your local appointments to remote calendar {0}", SelectedCalendar.Name);
                    foreach (var item in srcAppointmentItems)
                    {
                        AppointmentItem srcAppointment;
                        try
                        {
                            srcAppointment = item as AppointmentItem;
                        }
                        catch
                        {
                            srcAppointment = null;
                        }
                        if (srcAppointment == null)
                        {
                            Marshal.ReleaseComObject(item);
                            continue;
                        }

                        try
                        {
                            currentSubject = srcAppointment.Subject;

                            operationChain = "Checking source end date; ";
                            // skip non-recurring appintments which ended more than 90 days ago
                            if (!srcAppointment.IsRecurring && srcAppointment.End < DateTime.Now - TimeSpan.FromDays(30))
                            {
                                oldAppointments++;
                                continue;
                            }
                            // skip recurring appintments with last occurance more than 90 days ago
                            if (srcAppointment.IsRecurring)
                            {
                                var srcPattern = srcAppointment.GetRecurrencePattern();
                                if (!srcPattern.NoEndDate && srcPattern.PatternEndDate < DateTime.Now - TimeSpan.FromDays(30))
                                {
                                    oldAppointments++;
                                    continue;
                                }
                            }

                            IEvent dstAppointment = null;
                            operationChain += "Checking if target appointment alreay exists; ";
                            if (targetItems.ContainsKey(srcAppointment.GlobalAppointmentID))
                            {
                                dstAppointment = targetItems[srcAppointment.GlobalAppointmentID];
                                targetItems.Remove(srcAppointment.GlobalAppointmentID);
                                string lastSyncTime = GetLastSyncTimeStamp(srcAppointment);
                                if (lastSyncTime != null)
                                {
                                    // skip appointments which did not changed since last sync
                                    if (lastSyncTime == srcAppointment.LastModificationTime.ToString("O"))
                                    {
                                        unchangedAppointments++;
                                        Events.Add(new EventModel(dstAppointment){IsSynchronized = true});
                                        continue;
                                    }
                                }
                                // target appointments for which IsRecurring IsAllDay changed are deleted
                                // and we create a fresh one.
                                if (IsRecurring(dstAppointment) || srcAppointment.IsRecurring || srcAppointment.AllDayEvent != dstAppointment.IsAllDay)
                                {
                                    await dstAppointment.DeleteAsync();
                                    dstAppointment = null;
                                }
                            }

                            // this will indicate if we create a new appointment in the target folder
                            bool dstAppointmentIsNew = false;
                            try
                            {
                                WriteMessageLog("  Syncing [{0}] ", srcAppointment.Subject);

                                if (dstAppointment == null)
                                {
                                    dstAppointmentIsNew = true;
                                    operationChain += "Creating a new target appointment; ";
                                    dstAppointment = new Event();

                                    if (srcAppointment.AllDayEvent)
                                    {
                                        operationChain += "AllDayEvent; ";
                                        dstAppointment.IsAllDay = true;
                                    }
                                }

                                operationChain += "Updating Subject; ";
                                dstAppointment.Subject = srcAppointment.Subject;
                                operationChain += "Updating Location; ";
                                dstAppointment.Location = new Location {DisplayName = srcAppointment.Location};
                                operationChain += "Updating BusyStatus; ";
                                dstAppointment.ShowAs = GetFreeBusyStatus(srcAppointment.BusyStatus);

                                if (!srcAppointment.IsRecurring)
                                {
                                    WriteMessageLog("  on {0}", srcAppointment.Start);

                                    operationChain += "Not Recurring; ";

                                    operationChain += "Updating Start; ";
                                    dstAppointment.Start = CreateDateTimeTimeZone(srcAppointment.Start);

                                    operationChain += "Updating Duration; ";
                                    dstAppointment.End = CreateDateTimeTimeZone(srcAppointment.Start + TimeSpan.FromMinutes(srcAppointment.Duration));

                                    operationChain += "Updating originalLastUpdate; ";
                                    operationChain += "Saving; ";
                                    if (dstAppointmentIsNew)
                                    {
                                        await remoteCalendarEvents.AddEventAsync(dstAppointment);
                                        dstAppointmentItems.Add(dstAppointment);
                                        AddToMappingTable(srcAppointment, dstAppointment);
                                    }
                                    else
                                    {
                                        await dstAppointment.UpdateAsync();
                                    }
                                    Events.Add(new EventModel(dstAppointment) { IsSynchronized = true });
                                }
                                else // IsRecurring
                                {
                                    operationChain += "Recurring; ";
                                    var srcPattern = srcAppointment.GetRecurrencePattern();
                                    var dstRecurrence = dstAppointment.Recurrence ?? (dstAppointment.Recurrence = new PatternedRecurrence
                                    {
                                        Pattern = new RecurrencePattern(),
                                        Range = new RecurrenceRange()
                                    });

                                    operationChain += "Updating RecurreneType; ";
                                    dstRecurrence.Pattern.Type = GetPatternType(srcPattern.RecurrenceType);

                                    operationChain += "Updating StartTime; ";
                                    dstAppointment.Start = CreateDateTimeTimeZone(srcPattern.StartTime);
                                    operationChain += "Updating Duration; ";
                                    dstAppointment.End = CreateDateTimeTimeZone(srcPattern.EndTime);
                                    WriteMessageLog("  recurring {0} at {1} ", dstRecurrence.Pattern.Type, dstAppointment.Start.DateTime.Substring(11, 8));

                                    switch (srcPattern.RecurrenceType)
                                    {
                                        case OlRecurrenceType.olRecursDaily:
                                            break;
                                        case OlRecurrenceType.olRecursWeekly:
                                            dstRecurrence.Pattern.DaysOfWeek = CreateDaysOfWeekList(srcPattern.DayOfWeekMask);
                                            break;
                                        case OlRecurrenceType.olRecursMonthly:
                                            dstRecurrence.Pattern.DayOfMonth = srcPattern.DayOfMonth;
                                            break;
                                        case OlRecurrenceType.olRecursMonthNth:
                                            // Example: every second tuesday of the month, every 2 months would be:
                                            // Index will be 1 (first == 0, second == 1, ...)
                                            // DaysOfWeek will be tuesday
                                            // Interval will be 2 (every two months), covered below the switch for all cases.
                                            dstRecurrence.Pattern.Index = GetWeekIndex(srcPattern.Instance);
                                            dstRecurrence.Pattern.DaysOfWeek = CreateDaysOfWeekList(srcPattern.DayOfWeekMask);
                                            break;
                                        case OlRecurrenceType.olRecursYearly:
                                            dstRecurrence.Pattern.DayOfMonth = srcPattern.DayOfMonth;
                                            dstRecurrence.Pattern.Month = srcPattern.MonthOfYear;
                                            break;
                                        case OlRecurrenceType.olRecursYearNth:
                                            // example: every 2nd tuesday of january, every 2 years would be:
                                            // index = 1 (0 is first, 1 is seccond, ...),
                                            // DaysOfWeek = Tuesday 
                                            // Month would be 1 (january is 1, december is 12)
                                            // interval will be 2, covered below the else.
                                            dstRecurrence.Pattern.Index = GetWeekIndex(srcPattern.Instance);
                                            dstRecurrence.Pattern.DaysOfWeek = CreateDaysOfWeekList(srcPattern.DayOfWeekMask);
                                            dstRecurrence.Pattern.Month = srcPattern.MonthOfYear;
                                            break;

                                    }

                                    dstRecurrence.Range.StartDate = srcPattern.PatternStartDate.ToString("yyyy-MM-dd");
                                    dstRecurrence.Pattern.Interval = srcPattern.Interval > 0 ? srcPattern.Interval : 1;
                                    if (srcPattern.NoEndDate)
                                    {
                                        dstRecurrence.Range.Type = RecurrenceRangeType.NoEnd;
                                    }
                                    else if (srcPattern.Occurrences >= 0)
                                    {
                                        dstRecurrence.Range.Type = RecurrenceRangeType.Numbered;
                                        dstRecurrence.Range.NumberOfOccurrences = srcPattern.Occurrences;
                                    }
                                    else
                                    {
                                        dstRecurrence.Range.Type = RecurrenceRangeType.EndDate;
                                        dstRecurrence.Range.EndDate = srcPattern.PatternEndDate.ToString("yyyy-MM-dd");
                                    }

                                    operationChain += "Saving; ";
                                    if (dstAppointmentIsNew)
                                    {
                                        await remoteCalendarEvents.AddEventAsync(dstAppointment);
                                        dstAppointmentItems.Add(dstAppointment);
                                        AddToMappingTable(srcAppointment, dstAppointment);
                                    }
                                    else
                                    {
                                        await dstAppointment.UpdateAsync();
                                    }
                                    Events.Add(new EventModel(dstAppointment) { IsSynchronized = true });
                                    var srcExceptions = srcPattern.Exceptions;

                                    try
                                    {
                                        if (srcExceptions == null || srcExceptions.Count <= 0)
                                        {
                                            WriteMessageLog("    This recurring appointment has no exceptions.");
                                        }
                                        else
                                        {
                                            var intervalStart = DateTime.Now - TimeSpan.FromDays(30);
                                            var intervalEnd = DateTime.Now + TimeSpan.FromDays(90);
                                            
                                            WriteMessageLog("    Syncing exceptions between {0} and {1} for this recurring appointment:", intervalStart.ToShortDateString(), intervalEnd.ToShortDateString());

                                            var eventCollection = await remoteCalendarEvents[dstAppointment.Id].GetInstances(
                                                new DateTimeOffset(intervalStart), new DateTimeOffset(intervalEnd)).ExecuteAsync();

                                            List<IEvent> remoteInstances = await GetEventInstancesAsync(eventCollection);

                                            foreach (Microsoft.Office.Interop.Outlook.Exception srcException in srcExceptions)
                                            {
                                                var originalDate = new DateTime(srcException.OriginalDate.Year, srcException.OriginalDate.Month,
                                                    srcException.OriginalDate.Day).ToString("yyyy-MM-dd");

                                                AppointmentItem srcExceptionItem = null;

                                                try
                                                {
                                                    var dstExceptionItem = remoteInstances.FirstOrDefault(ri => ri.Start.DateTime.Substring(0, 10) == originalDate);
                                                    if (dstExceptionItem == null)
                                                    {
                                                        // out of interval, ignore this
                                                        continue;
                                                    }
                                                    // we need to re-fetch it becasue changes in this loop on one of the items might affect the other items.
                                                    // But the Id is stable!
                                                    dstExceptionItem = await remoteCalendarEvents[dstExceptionItem.Id].ExecuteAsync();
                                                    if (srcException.Deleted)
                                                    {
                                                        await dstExceptionItem.DeleteAsync();
                                                        WriteMessageLog("      On {0}: occurence cancelled.", originalDate);
                                                    }
                                                    else
                                                    {
                                                        srcExceptionItem = srcException.AppointmentItem;
                                                        WriteMessageLog("      On {0}: shifting occurence to {1}.", originalDate, srcExceptionItem.Start);

                                                        dstExceptionItem.SeriesMasterId = dstAppointment.Id;
                                                        dstExceptionItem.Type = EventType.Exception;
                                                        dstExceptionItem.Subject = srcExceptionItem.Subject;
                                                        dstExceptionItem.Location = new Location {DisplayName = srcExceptionItem.Location};
                                                        dstExceptionItem.Start = CreateDateTimeTimeZone(srcExceptionItem.Start);
                                                        dstExceptionItem.End = CreateDateTimeTimeZone(srcExceptionItem.End);
                                                        await dstExceptionItem.UpdateAsync();
                                                    }
                                                }

                                                finally
                                                {
                                                    if (srcExceptionItem != null)
                                                    {
                                                        Marshal.ReleaseComObject(srcExceptionItem);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        Marshal.ReleaseComObject(srcExceptions);
                                    }
                                }
                                successfullyUpdated++;
                                SetLastSyncTimeStamp(srcAppointment);
                            }
                            catch (Exception ex)
                            {
                                if (dstAppointmentIsNew)
                                {
                                    WriteMessageLog("  ERROR: Could not create appointment [{0}] in target calendar.", srcAppointment.Subject);
                                    WriteMessageLog("  Chain of performed operations: {0}", operationChain);
                                    WriteMessageLog("  {0}", ex.ToString());
                                    WriteMessageLog("");
                                }
                                else
                                {
                                    WriteMessageLog("  ERROR: Could not sync appointment [{0}].", srcAppointment.Subject);
                                }
                                errorUpdated++;
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteMessageLog("ERROR syncing [{0}]. The message below might help us understanding what happened. Sorry.", currentSubject);
                            WriteMessageLog("Chain of performed operations: {0}", operationChain);
                            WriteMessageLog("{0}", ex.ToString());
                            WriteMessageLog("");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(srcAppointment);
                        }
                    }

                    SaveMappingTable();

                    int deletedItems = await DeleteAppointments(targetItems);

                    deletedUpdates += deletedItems;
                    errorUpdated += (targetItems.Count - deletedItems);

                    WriteMessageLog("Sync done");
                    WriteMessageLog("{0} appointments updated.", successfullyUpdated);
                    WriteMessageLog("{0} appointments deleted.", deletedUpdates);
                    WriteMessageLog("{0} appointments failed to be updated.", errorUpdated);
                    WriteMessageLog("{0} appointments not synced because they ended more than 30 days ago.", oldAppointments);
                    WriteMessageLog("{0} appointments not synced because they did not changed since their last sync.", unchangedAppointments);
                    WriteMessageLog("");
                }
            }
            catch (Exception ex)
            {
                WriteMessageLog("Error during synchronization: {0}", ex);
                Error = "Error during synchronization, see log file";
            }
            finally
            {
                if (srcAppointmentItems != null)
                {
                    Marshal.ReleaseComObject(srcAppointmentItems);
                }
                RunningAsyncOperations--;
                if (IsAutoSyncEnabled)
                {
                    myTimer.Interval = TimeSpan.FromMinutes(AutoSyncInterval);
                    myTimer.Start();
                }
            }
        }

        async private Task<List<IEvent>> GetEventInstancesAsync(IPagedCollection<IEvent> eventCollection)
        {
            var result = new List<IEvent>();
            while (eventCollection != null)
            {
                result.AddRange(eventCollection.CurrentPage);
                eventCollection = eventCollection.MorePagesAvailable ? await eventCollection.GetNextPageAsync() : null;
            }
            return result;
        }

        async private Task<int> DeleteAppointments(Dictionary<string, IEvent> targetItems)
        {
            var deletedItems = 0;
            foreach (var item in targetItems)
            {
                WriteMessageLog("  Deleting [{0}]", item.Value.Subject);
                try
                {
                    myMappingTable.Remove(item.Key);
                    await item.Value.DeleteAsync();
                    deletedItems++;
                }
                catch (Exception ex)
                {
                    WriteMessageLog("  ERROR: Could not delete appointment [{0}]: {1}", item.Value.Subject, ex.Message);
                }                
            }
            return deletedItems;
        }

        private WeekIndex GetWeekIndex(int p)
        {
            return (WeekIndex) (p-1);
        }

        private IList<DayOfWeek> CreateDaysOfWeekList(OlDaysOfWeek olDaysOfWeekMask)
        {
            int olDaysOfWeek = (int) olDaysOfWeekMask;

            var result = new List<DayOfWeek>();
            if ((olDaysOfWeek & (int)OlDaysOfWeek.olMonday) != 0)
            {
                result.Add(DayOfWeek.Monday);
            }
            if ((olDaysOfWeek & (int)OlDaysOfWeek.olTuesday) != 0)
            {
                result.Add(DayOfWeek.Tuesday);
            }
            if ((olDaysOfWeek & (int)OlDaysOfWeek.olWednesday) != 0)
            {
                result.Add(DayOfWeek.Wednesday);
            }
            if ((olDaysOfWeek & (int)OlDaysOfWeek.olThursday) != 0)
            {
                result.Add(DayOfWeek.Thursday);
            }
            if ((olDaysOfWeek & (int)OlDaysOfWeek.olFriday) != 0)
            {
                result.Add(DayOfWeek.Friday);
            }
            if ((olDaysOfWeek & (int)OlDaysOfWeek.olSaturday) != 0)
            {
                result.Add(DayOfWeek.Saturday);
            }
            if ((olDaysOfWeek & (int)OlDaysOfWeek.olSunday) != 0)
            {
                result.Add(DayOfWeek.Sunday);
            }
            return result;
        }

        private RecurrencePatternType GetPatternType(OlRecurrenceType olRecurrenceType)
        {
            switch (olRecurrenceType)
            {
                case OlRecurrenceType.olRecursMonthNth:
                    return RecurrencePatternType.RelativeMonthly;
                case OlRecurrenceType.olRecursMonthly:
                    return RecurrencePatternType.AbsoluteMonthly;
                case OlRecurrenceType.olRecursWeekly:
                    return RecurrencePatternType.Weekly;
                case OlRecurrenceType.olRecursYearNth:
                    return RecurrencePatternType.RelativeYearly;
                case OlRecurrenceType.olRecursYearly:
                    return RecurrencePatternType.AbsoluteYearly;
                default:
                    return RecurrencePatternType.Daily;
            }
        }

        private DateTimeTimeZone CreateDateTimeTimeZone(DateTime dateTime)
        {
            return new DateTimeTimeZone
            {
                DateTime = dateTime.ToString("O"),
                TimeZone = TimeZoneInfo.Local.Id
            };
        }

        private FreeBusyStatus GetFreeBusyStatus(OlBusyStatus olBusyStatus)
        {
            switch (olBusyStatus)
            {
                case OlBusyStatus.olBusy:
                    return FreeBusyStatus.Busy;
                case OlBusyStatus.olFree:
                    return FreeBusyStatus.Free;
                case OlBusyStatus.olOutOfOffice:
                    return FreeBusyStatus.Oof;
                case OlBusyStatus.olTentative:
                    return FreeBusyStatus.Tentative;
                default:
                    return FreeBusyStatus.Unknown;
            }
        }

        private string GetLastSyncTimeStamp(AppointmentItem srcAppointment)
        {
            //return dstAppointment.UserProperties["originalLastUpdate"];
            if (myMappingTable.ContainsKey(srcAppointment.GlobalAppointmentID))
            {
                var identificationInfo = myMappingTable[srcAppointment.GlobalAppointmentID];
                if (identificationInfo != null)
                {
                    return identificationInfo.LastSyncTimeStamp;
                }
            }
            return null;
        }

        private void SetLastSyncTimeStamp(AppointmentItem srcAppointment)
        {
            myMappingTable[srcAppointment.GlobalAppointmentID].LastSyncTimeStamp = srcAppointment.LastModificationTime.ToString("O");
        }

        private void AddToMappingTable(AppointmentItem srcAppointment, IEvent dstAppointment)
        {
            if (myMappingTable.ContainsKey(srcAppointment.GlobalAppointmentID)) myMappingTable.Remove(srcAppointment.GlobalAppointmentID);
            myMappingTable.Add(srcAppointment.GlobalAppointmentID,
                new EventIdentificationInfo
                {
                    Id = dstAppointment.Id,
                    LastSyncTimeStamp = srcAppointment.LastModificationTime.ToString("O")
                }
                );
        }

        private void RemoveFromMappingTable(IEvent e)
        {
            var key = myMappingTable.Where(kvp => kvp.Value.Id == e.Id).Select(kvp => kvp.Key).FirstOrDefault();
            if (key != null)
            {
                myMappingTable.Remove(key);
            }
        }

        private static bool IsRecurring(IEvent dstAppointment)
        {
            return dstAppointment.Recurrence != null && dstAppointment.Type != EventType.SingleInstance;
        }

        private Dictionary<string, IEvent> GetAllSyncedEvents(IList<IEvent> dstAppointmentItems)
        {

            var targetItems = new Dictionary<string, IEvent>();
            var keysToBeRemoved = new List<string>();
            foreach (var key in myMappingTable.Keys)
            {
                var dst = dstAppointmentItems.FirstOrDefault(e => e.Id == myMappingTable[key].Id);
                if(dst != null) targetItems.Add(key, dst);
                else keysToBeRemoved.Add(key);
            }
            foreach (var key in keysToBeRemoved)
            {
                myMappingTable.Remove(key);
            }
            return targetItems;
        }

        async private Task DeleteEventAsync(EventModel e)
        {
            if (e != null)
            {
                var pos = Events.IndexOf(e);
                SelectedEvent = null;
                RemoveFromMappingTable(e.Event);
                await e.Event.DeleteAsync();
                Events.Remove(e);
                if (pos >= Events.Count)
                {
                    pos--;
                }
                if (pos >= 0)
                {
                    SelectedEvent = Events[pos];
                }
            }
        }
        async private Task DeleteAllEventsAsync()
        {
            RunningAsyncOperations++;
            try
            {
                var events = Events.ToList();
                foreach (var e in events)
                {
                    await DeleteEventAsync(e);
                }
            }
            catch (Exception exception)
            {
                WriteMessageLog("Could not delete all events: {0}", exception);
            }
            finally
            {
                RunningAsyncOperations--;
            }
        }

        public void Cleanup()
        {
            Calendars.Clear();
            SelectedCalendar = null;
        }

        public void Close()
        {
            if (myMessageLog != null)
            {
                myMessageLog.Close();
                myMessageLog.Dispose();
            }
        }

        async private void AutologinAsync()
        {
            if (String.IsNullOrEmpty(Properties.Settings.Default.refresh_token))
            {
                Error = "You are not logged in";
                return;
            }
            RunningAsyncOperations++;
            try
            {
                CreateOutlookClient();
                IsLoggedIn = true;
                await GetUserDataAsync();
                await GetUserCalendarsAsync();
            }
            catch (Exception ex)
            {
                if (!IsLoggedIn)
                {
                    Properties.Settings.Default.refresh_token = null;
                    Properties.Settings.Default.Save();
                }
                WriteMessageLog("Could not log you in: {0}", ex);
                LogApiMigrationError();
            }
            finally
            {
                RunningAsyncOperations--;
            }
        }

        private void CreateOutlookClient()
        {
            myOutlookServicesClient = new OutlookServicesClient(new Uri("https://outlook.office.com/api/v2.0"), OAuthHelper.GetAccessTokenTaskAsync);
        }

        async public void LoginAsync(string authCode)
        {
            if (String.IsNullOrEmpty(authCode))
            {
                throw new ArgumentNullException("authCode");
            }
            RunningAsyncOperations++;
            try
            {
                await OAuthHelper.AuthorizeTaskAsync(authCode);
                AutologinAsync();
            }
            catch(Exception ex)
            {
                if (!IsLoggedIn)
                {
                    Properties.Settings.Default.refresh_token = null;
                    Properties.Settings.Default.Save();                        
                }
                WriteMessageLog("Could not log you in: {0}", ex);
                Error = "Login Failed";
            }
            finally
            {
                RunningAsyncOperations--;
            }
        }

        private void LogApiMigrationError()
        {
            WriteMessageLog("");
            WriteMessageLog("Lumis Calendar sync is using a new Microsoft API, and not all Microsoft accounts are enabled for this API.");
            WriteMessageLog("The most probable cause for the error above is that your account has not yet been enabled.");
            WriteMessageLog("It is not clear how log it will take until your account will be migrated.");
            WriteMessageLog(
                "Currently, the only workaround is for you to create a new Microsoft account and use it, as all new accounts have the new API automatically enabled.");
            WriteMessageLog(
                "Alternatively, continue to use the previous version of Lumis Calendar Sync and try this new version perdiodically until your account gets migrated.");
            Error = "We cannot access your Outlook.com Calendars. Your account might not have been migrated yet.";
        }

        async public void LogoutAsync()
        {
            SelectedEvent = null;
            Events.Clear();
            SelectedCalendar = null;
            Calendars.Clear();

            RunningAsyncOperations++;
            try
            {
                await OAuthHelper.LogoutTaskAsync();
            }
            catch (Exception exception)
            {
                WriteMessageLog("Error logging you out: {0}", exception);
            }
            finally
            {
                RunningAsyncOperations--;
                Properties.Settings.Default.refresh_token = null;
                Properties.Settings.Default.Save();
                IsLoggedIn = false;
                User = null;
                Cleanup();
            }
        }

        async private Task GetUserDataAsync()
        {
            User = await myOutlookServicesClient.Me.ExecuteAsync();
        }

        async private Task GetUserCalendarsAsync()
        {
            var cals = await myOutlookServicesClient.Me.Calendars.ExecuteAsync();
            while(cals != null)
            {
                foreach (var cal in cals.CurrentPage)
                {
                    Calendars.Add(cal);
                    if (cal.Id == Properties.Settings.Default.RemoteCaleandarId)
                    {
                        SelectedCalendar = cal;
                    }
                }
                cals = cals.MorePagesAvailable ? await cals.GetNextPageAsync() : null;
            }
        }

        async private void PopulateEventsAsync(string calendarId)
        {
            RunningAsyncOperations++;
            try
            {
                Events.Clear();
                var newEvents = await GetCalendarEventsAsync(calendarId);
                foreach (var e in newEvents)
                {
                    var em = new EventModel(e);
                    if (IsEventInMappingTable(e))
                    {
                        em.IsSynchronized = true;
                    }
                    Events.Add(em);
                }
                
                if (IsAutoSyncEnabled)
                {
                    myTimer.Interval = TimeSpan.FromMinutes(1);
                    myTimer.Start();
                }
            }
            catch (Exception ex)
            {
                WriteMessageLog("Error in retrieving the remote events: {0}", ex);
            }
            finally
            {
                RunningAsyncOperations--;
            }
        }

        private bool IsEventInMappingTable(IEvent e)
        {
            return myMappingTable.Values.Any(entry => entry.Id == e.Id);
        }

        async private Task<IList<IEvent>> GetCalendarEventsAsync(string calendarid)
        {
            RunningAsyncOperations++;
            var result = new List<IEvent>();

            var evts = await myOutlookServicesClient.Me.Calendars[calendarid].Events.ExecuteAsync();
            while(evts != null)
            {
                result.AddRange(evts.CurrentPage);
                evts = evts.MorePagesAvailable ? await evts.GetNextPageAsync() : null;
            }
            RunningAsyncOperations--;

            return result;
        }

        private void CreateStartupFolderShortcut()
        {
            var wshShell = new WshShellClass();
            string startUpFolderPath =
              Environment.GetFolderPath(Environment.SpecialFolder.Startup);

            // Create the shortcut
            var shortcut = (IWshShortcut)wshShell.CreateShortcut(
                startUpFolderPath + "\\" +
                "LumisCalendarSync" + ".lnk");

            shortcut.TargetPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            shortcut.WorkingDirectory = System.Reflection.Assembly.GetExecutingAssembly().Location;
            shortcut.Arguments = "/Minimized";
            shortcut.Description = "Lumis Calendar Sync";
            shortcut.Save();
        }

        private static void DeleteStartupFolderShortcut()
        {
            var startUpFolderPath =
              Environment.GetFolderPath(Environment.SpecialFolder.Startup);

            var di = new DirectoryInfo(startUpFolderPath);
            var files = di.GetFiles("*.lnk");

            var fileInfo = files.FirstOrDefault(fi => fi.Name == "LumisCalendarSync.lnk");
            if(fileInfo != null)
            {
                File.Delete(fileInfo.FullName);
            }
        }

        private static bool IsStartupShortcutSaved()
        {
            var startUpFolderPath =
              Environment.GetFolderPath(Environment.SpecialFolder.Startup);

            var di = new DirectoryInfo(startUpFolderPath);
            var files = di.GetFiles("*.lnk");

            var fileInfo = files.FirstOrDefault(fi => fi.Name == "LumisCalendarSync.lnk");
            return fileInfo != null;
        }

        public readonly OAuthHelper OAuthHelper;
        private IOutlookServicesClient myOutlookServicesClient;
        private Dictionary<string, EventIdentificationInfo> myMappingTable;
        private StreamWriter myMessageLog;
        private readonly DispatcherTimer myTimer;
        private string myAppDataFolder;
    }
}

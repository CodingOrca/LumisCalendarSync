using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Threading;
using IWshRuntimeLibrary;
using LumisCalendarSync.Model;
using LumisCalendarSync.Properties;

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
        // This version is shown to the user in the UI and shall be the same as the version set in the MSI.
        // We change this whenever we publish a new msi version.
        public string CurrentAppVersion
        {
            get { return "2.2.0.0"; }
        }
        
        // When we change the list of synced attributes, we change CurrentDataVersion to force a sync of all appointments 
        // at the first sync with the new app version. 
        // Best Practice: if we change some code (fix or new feature) which needs a full sync, set this to the same value as in the MSI.
        // But not needing to set it vor every new MSI version: if no sync must be forced, don't change it.
        public string CurrentDataVersion
        {
            get { return "2.2.0.0"; }
        }


        public MainViewModel()
        {
            OAuthHelper = new OAuthHelper();
            Calendars = new ObservableCollection<ICalendar>();
            Events = new ObservableCollection<EventModel>();
            LogEntries = new ObservableCollection<string>();

            if (CurrentDataVersion != Settings.Default.AppVersion)
            {
                Settings.Default.AppVersion = CurrentDataVersion;
                Settings.Default.ForceNextSync = true;
                Settings.Default.Save();
            }
            
            myTimer = new DispatcherTimer();
            myTimer.Tick += Timer_Tick;
            myTimer.Interval = TimeSpan.FromMinutes(Settings.Default.AutoSyncInterval);

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
            // delete the file if it exceeds 100 KB
            if (myMessageLog != null && myMessageLog.BaseStream.Length > 100 * 1024)
            {
                myMessageLog.Close();
                File.Delete(LogFileName);
                myMessageLog = null;
            }
            if (myMessageLog == null)
            {
                try
                {
                    myMessageLog = new StreamWriter(LogFileName, append:true) { AutoFlush = true };
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
            var mappingFile = Path.Combine(myAppDataFolder, String.Format("{0}-{1}.mapping", User.EmailAddress, SelectedCalendar.Name));
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
            // Legacy: versions up to 2.0.8 where using a hashcode of the calendar id for the file name, which is unstable
            // (different between debug and release, can be changed from .net version)
            // So we move it to a mre reliable filename:
            var legacyMappingFile = Path.Combine(myAppDataFolder, String.Format("Mapping-{0}.dat", SelectedCalendar.Id.GetHashCode()));
            var mappingFile = Path.Combine(myAppDataFolder, String.Format("{0}-{1}.mapping", User.EmailAddress, SelectedCalendar.Name));

            if (File.Exists(legacyMappingFile))
            {
                if(!File.Exists(mappingFile)) File.Move(legacyMappingFile, mappingFile );
                else File.Delete(legacyMappingFile);
            }
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
                    Settings.Default.RemoteCaleandarId = SelectedCalendar.Id;
                    Settings.Default.Save();
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
            get { return Settings.Default.IsAutoSyncEnabled; }
            set
            {
                Settings.Default.IsAutoSyncEnabled = value;
                Settings.Default.Save();
                RaisePropertyChanged("IsAutoSyncEnabled");
                RaisePropertyChanged("CanChangeCalendar");
                myTimer.IsEnabled = value;
            }
        }

        public int AutoSyncInterval
        {
            get { return Settings.Default.AutoSyncInterval; }
            set
            {
                if(value < 10 ) throw new Exception("Provide a value in minutes > 10");
                Settings.Default.AutoSyncInterval = value;
                Settings.Default.Save();
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

        public bool SkipOldAppointments
        {
            get { return Settings.Default.SkipOldAppointments; }
            set
            {
                Settings.Default.SkipOldAppointments = value;
                Settings.Default.Save();
                RaisePropertyChanged("SkipOldAppointments");
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
            if (String.IsNullOrEmpty(Settings.Default.RemoteCaleandarId))
            {
                Error = "No remote calendar selected";
                return;
            }

            int unchangedAppointments = 0;
            int skippedAppointments = 0;
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
                        WriteMessageLog("Outlook is not running, cannot sync.");
                        Error = "Outlook is not running, cannot sync.";
                        return;
                    }

                    var remoteCalendarEvents = myOutlookServicesClient.Me.Calendars[Settings.Default.RemoteCaleandarId].Events;
                    
                    Events.Clear();
                    LogEntries.Clear();
                    WriteMessageLog("Starting syncing your local appointments to remote calendar [{0}] on account [{1}].", SelectedCalendar.Name, User.EmailAddress);

                    var dstAppointmentItems = await GetCalendarEventsAsync(Settings.Default.RemoteCaleandarId);

                    var targetItems = GetAllSyncedEvents(dstAppointmentItems);

                    foreach (var e in dstAppointmentItems)
                    {
                        if (targetItems.Values.Contains(e)) continue;
                        Events.Add(new EventModel(e) { IsSynchronized = false });
                    }

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
                            operationChain = "";

                            if (SkipOldAppointments && IsAppointmentOld(srcAppointment))
                            {
                                WriteMessageLog("Skipping [{0}] as it ends more than 30 days ago.", srcAppointment.Subject);
                                skippedAppointments++;
                                continue;
                            }

                            currentSubject = srcAppointment.Subject;
                            IEvent dstAppointment = null;
                            operationChain += "Checking if target appointment alreay exists; ";
                            string reasonForSync = "New Appointment";
                            if (targetItems.ContainsKey(srcAppointment.GlobalAppointmentID))
                            {
                                dstAppointment = targetItems[srcAppointment.GlobalAppointmentID];
                                targetItems.Remove(srcAppointment.GlobalAppointmentID);
                                string lastSyncTime = GetLastSyncTimeStamp(srcAppointment);
                                // skip appointments which did not changed since last sync
                                if (lastSyncTime == srcAppointment.LastModificationTime.ToString("O") && !Settings.Default.ForceNextSync)
                                {
                                    unchangedAppointments++;
                                    Events.Add(new EventModel(dstAppointment){IsSynchronized = true});
                                    continue;
                                }

                                reasonForSync = Settings.Default.ForceNextSync?
                                                    "Application Updated" : HasAppointmentInformationChanged(dstAppointment, srcAppointment);

                                // if nothing changed and the appointment is not recurring, no sync is needed, just update the last synced time stamp:
                                if (reasonForSync == null && !srcAppointment.IsRecurring )
                                {
                                    SetLastSyncTimeStamp(srcAppointment);
                                    unchangedAppointments++;
                                    Events.Add(new EventModel(dstAppointment) { IsSynchronized = true });
                                    continue;
                                }

                                // we get here if something has changed OR the appointment is recurring.

                                // changed appointments which are Recurring or the IsAllDay attribute changed, are deleted since we cannot update them correctly.
                                // (we create a new target appoitment for them later on!)
                                if (reasonForSync != null && (srcAppointment.IsRecurring || srcAppointment.AllDayEvent != dstAppointment.IsAllDay))
                                {
                                    RemoveFromMappingTable(dstAppointment);
                                    await dstAppointment.DeleteAsync();
                                    dstAppointment = null;
                                }
                            }

                            // this indicates if we create a new appointment in the remote calendar
                            bool dstAppointmentIsNew = (dstAppointment == null);
                            try
                            {
                                WriteMessageLog("  Syncing [{0}]: {1}.", srcAppointment.Subject, reasonForSync);

                                if (dstAppointment == null)
                                {
                                    operationChain += "Creating a new target appointment; ";
                                    dstAppointment = new Event();

                                    if (srcAppointment.AllDayEvent)
                                    {
                                        operationChain += "AllDayEvent; ";
                                        dstAppointment.IsAllDay = true;
                                    }
                                }

                                if (reasonForSync != null)
                                {
                                    operationChain += "Updating Subject; ";
                                    dstAppointment.Subject = srcAppointment.Subject;
                                    operationChain += "Updating Location; ";
                                    dstAppointment.Location = new Location {DisplayName = srcAppointment.Location};
                                    operationChain += "Updating BusyStatus; ";
                                    dstAppointment.ShowAs = GetFreeBusyStatus(srcAppointment.BusyStatus);
                                    dstAppointment.IsReminderOn = srcAppointment.ReminderSet;

                                    if (srcAppointment.ReminderSet)
                                    {
                                        dstAppointment.ReminderMinutesBeforeStart = srcAppointment.ReminderMinutesBeforeStart;
                                    }
                                }

                                // Non-recurring appointment:
                                if (!srcAppointment.IsRecurring)
                                {
                                    WriteMessageLog("  on {0}.", srcAppointment.Start);

                                    operationChain += "Not Recurring; ";

                                    operationChain += "Updating Start and End; ";
                                    if (srcAppointment.AllDayEvent)
                                    {
                                        dstAppointment.Start = CreateDateTimeTimeZone(srcAppointment.Start.Date);
                                        dstAppointment.End = CreateDateTimeTimeZone(srcAppointment.End.Date);
                                    }
                                    else
                                    {
                                        dstAppointment.Start = CreateDateTimeTimeZone(srcAppointment.Start);
                                        dstAppointment.End = CreateDateTimeTimeZone(srcAppointment.End);
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
                                }
                                else // Recurring appointment
                                {
                                    operationChain += "Recurring; ";
                                    var srcPattern = srcAppointment.GetRecurrencePattern();

                                    // if the recurrenct appointment changed, we have created a new one; else, we need no change for the master event.
                                    if (dstAppointmentIsNew)
                                    {
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
                                        WriteMessageLog("  recurring {0} at {1}.", dstRecurrence.Pattern.Type, dstAppointment.Start.DateTime.Substring(11, 8));

                                        UpdateDestinationPattern(srcPattern, dstRecurrence);

                                        operationChain += "Saving; ";
                                        await remoteCalendarEvents.AddEventAsync(dstAppointment);
                                        dstAppointmentItems.Add(dstAppointment);
                                        AddToMappingTable(srcAppointment, dstAppointment);
                                    }
                                    else
                                    {
                                        WriteMessageLog("  No change for the series master, no update needed.");
                                    }
                                    Events.Add(new EventModel(dstAppointment) { IsSynchronized = true });

                                    // Handling exceptions for the recurring appointment:
                                    var srcExceptions = srcPattern.Exceptions;
                                    try
                                    {
                                        if (srcExceptions.Count > 0)
                                        {
                                            WriteMessageLog("    Syncing exceptions for this series:");
                                        }
                                        var numberOfUnchangedExceptions = 0;
                                        foreach (Microsoft.Office.Interop.Outlook.Exception srcException in srcExceptions)
                                        {
                                            var srcExceptionItem = srcException.Deleted ? null : srcException.AppointmentItem;
                                            try
                                            {
                                                if (SkipOldAppointments)
                                                {
                                                    if (srcException.Deleted && (DateTime.Now.Date - srcException.OriginalDate.Date).TotalDays > 30)
                                                    {
                                                        numberOfUnchangedExceptions++;
                                                        continue;
                                                    }
                                                    if (srcExceptionItem != null && (DateTime.Now.Date - srcExceptionItem.End.Date).TotalDays > 30)
                                                    {
                                                        numberOfUnchangedExceptions++;
                                                        continue;
                                                    }
                                                }

                                                IEvent dstExceptionItem = null;
                                                var originalDate = srcException.OriginalDate.ToString("O").Substring(0, 10);
                                                if (myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.ContainsKey(originalDate))
                                                {
                                                    var lastChange = srcExceptionItem != null ? srcExceptionItem.LastModificationTime.ToString("O") : null;
                                                    if (lastChange == myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].LastSyncTimeStamp)
                                                    {
                                                        numberOfUnchangedExceptions++;
                                                        continue;
                                                    }
                                                    var remoteId = myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].Id;
                                                    if (remoteId != null)
                                                    {
                                                        dstExceptionItem = await remoteCalendarEvents[remoteId].ExecuteAsync();
                                                    }
                                                }
                                                else
                                                {
                                                    var intervalStart = srcException.OriginalDate - TimeSpan.FromDays(1);
                                                    var intervalEnd = srcException.OriginalDate + TimeSpan.FromDays(1);
                                                    var eventCollection = await remoteCalendarEvents[dstAppointment.Id].GetInstances(
                                                        new DateTimeOffset(intervalStart), new DateTimeOffset(intervalEnd)).ExecuteAsync();
                                                    var remoteInstances = await GetEventInstancesAsync(eventCollection);
                                                    dstExceptionItem = remoteInstances.FirstOrDefault(ri => GetLocalTime(ri.Start).ToString("O").Substring(0, 10) == originalDate);
                                                }
                                                if (dstExceptionItem == null)
                                                {
                                                    WriteMessageLog("      No remote instance found for local {1}exception on {0}.", originalDate,
                                                        srcException.Deleted ? "deleted " : "");
                                                    continue;
                                                }
                                                // we need to re-fetch it becasue changes in this loop on one of the items might affect the other items.
                                                // But the Id is stable!
                                                // dstExceptionItem = await remoteCalendarEvents[dstExceptionItem.Id].ExecuteAsync();
                                                if (srcException.Deleted)
                                                {
                                                    await dstExceptionItem.DeleteAsync();
                                                    WriteMessageLog("      On {0}: cancelled.", originalDate);
                                                }
                                                else
                                                {
                                                    dstExceptionItem.SeriesMasterId = dstAppointment.Id;
                                                    dstExceptionItem.Type = EventType.Exception;
                                                    dstExceptionItem.Subject = srcExceptionItem.Subject;
                                                    dstExceptionItem.Location = new Location {DisplayName = srcExceptionItem.Location};
                                                    dstExceptionItem.Start = CreateDateTimeTimeZone(srcExceptionItem.Start);
                                                    dstExceptionItem.End = CreateDateTimeTimeZone(srcExceptionItem.End);
                                                    dstExceptionItem.IsReminderOn = srcExceptionItem.ReminderSet;
                                                    if (srcExceptionItem.ReminderSet)
                                                    {
                                                        dstExceptionItem.ReminderMinutesBeforeStart = srcExceptionItem.ReminderMinutesBeforeStart;
                                                    }
                                                    await dstExceptionItem.UpdateAsync();
                                                    WriteMessageLog("      On {0}: updated, start at {1}.", originalDate, srcExceptionItem.Start.ToString("g", CultureInfo.CurrentCulture));

                                                    UpdateExceptionInMappingTable(srcAppointment.GlobalAppointmentID, originalDate, dstExceptionItem.Id);
                                                }
                                            }

                                            finally
                                            {
                                                if (srcExceptionItem != null)
                                                {
                                                    Marshal.ReleaseComObject(srcExceptionItem);
                                                }
                                                Marshal.ReleaseComObject(srcException);
                                            }
                                        }
                                        if (srcExceptions.Count > 0)
                                        {
                                            WriteMessageLog("    {0} Exceptions found, {1} needs no update or are too old.", srcExceptions.Count, numberOfUnchangedExceptions);
                                        }
                                        else
                                        {
                                            WriteMessageLog("    This series has no exceptions.");
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
                                    WriteMessageLog("  Chain of performed operations: {0}.", operationChain);
                                    WriteMessageLog("  {0}", ex.ToString());
                                    WriteMessageLog("");
                                }
                                else
                                {
                                    WriteMessageLog("  ERROR: Could not sync appointment [{0}].", srcAppointment.Subject);
                                    WriteMessageLog("  {0}", ex.ToString());
                                    WriteMessageLog("");
                                }
                                errorUpdated++;
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteMessageLog("ERROR syncing [{0}]. The message below might help us understanding what happened. Sorry.", currentSubject);
                            WriteMessageLog("Chain of performed operations: {0}.", operationChain);
                            WriteMessageLog("{0}", ex.ToString());
                            WriteMessageLog("");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(srcAppointment);
                        }
                    }

                    int deletedItems = await DeleteAppointments(targetItems);
                    SaveMappingTable();

                    deletedUpdates += deletedItems;
                    errorUpdated += (targetItems.Count - deletedItems);

                    WriteMessageLog("Sync done.");
                    if( successfullyUpdated != 0) WriteMessageLog("{0} appointments updated / created.", successfullyUpdated);
                    if (deletedUpdates != 0) WriteMessageLog("{0} appointments deleted.", deletedUpdates);
                    if (errorUpdated != 0) WriteMessageLog("{0} appointments failed to be updated.", errorUpdated);
                    if (unchangedAppointments != 0) WriteMessageLog("{0} appointments did not changed since their last sync.", unchangedAppointments);
                    if(skippedAppointments != 0) WriteMessageLog("{0} appointments not synced or deleted because they are older than 30 days.", skippedAppointments);
                    WriteMessageLog("");

                    Settings.Default.ForceNextSync = false;
                    Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                WriteMessageLog("Error during synchronization: {0}", ex);
                Error = "Error during synchronization, see log file.";
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

        private static bool IsAppointmentOld(AppointmentItem srcAppointment)
        {
            bool isToBeSkipped = false;
            if (srcAppointment.IsRecurring)
            {
                var pattern = srcAppointment.GetRecurrencePattern();
                isToBeSkipped = !pattern.NoEndDate &&
                                (DateTime.Now.Date - pattern.PatternEndDate).TotalDays > 30;
                Marshal.ReleaseComObject(pattern);
            }
            else
            {
                isToBeSkipped = (DateTime.Now.Date - srcAppointment.End.Date).TotalDays > 30;
            }
            return isToBeSkipped;
        }


        private static void UpdateDestinationPattern(Microsoft.Office.Interop.Outlook.RecurrencePattern srcPattern, PatternedRecurrence dstRecurrence)
        {
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
        }

        private string HasAppointmentInformationChanged(IEvent dstAppointment, AppointmentItem srcAppointment)
        {
            if (IsRecurring(dstAppointment) != srcAppointment.IsRecurring) return "Recurrence changed";
            if (dstAppointment.IsAllDay != srcAppointment.AllDayEvent) return "All Day changed";
            if (dstAppointment.Subject != srcAppointment.Subject) return "Subject Changed";
            if (dstAppointment.Location == null) return "Destination Location is empty";
            if (!String.IsNullOrEmpty(dstAppointment.Location.DisplayName) || !String.IsNullOrEmpty(srcAppointment.Location))
            {
                if (dstAppointment.Location.DisplayName != srcAppointment.Location)
                {
                    return "Location changed";
                }
            }
            if (dstAppointment.ShowAs != GetFreeBusyStatus(srcAppointment.BusyStatus)) return "FreeBusyStatus changed";
            if (dstAppointment.IsReminderOn != srcAppointment.ReminderSet) return "ReminderSet changed";
            if (srcAppointment.ReminderSet)
            {
                if (dstAppointment.ReminderMinutesBeforeStart != srcAppointment.ReminderMinutesBeforeStart) return "Reminder value changed";
            }

            if (!srcAppointment.IsRecurring)
            {
                if (!GetLocalTime(dstAppointment.Start).Equals(srcAppointment.Start)) return "Start changed";
                if (!GetLocalTime(dstAppointment.End).Equals(srcAppointment.End)) return "Duration changed";
                return null;
            }

            // from here on, we deal with a recurring appointment!

            var srcPattern = srcAppointment.GetRecurrencePattern();
            var dstRecurrence = dstAppointment.Recurrence;
            
            if (srcPattern == null) return "Source Recurrence Pattern is not set";
            if (dstRecurrence == null) return "Destination recurrence pattern is not set";
            if (dstRecurrence.Pattern.Type != GetPatternType(srcPattern.RecurrenceType)) return "RecurrenceType changed";
            if (!IsTimeIdentical(GetLocalTime(dstAppointment.Start), srcPattern.StartTime)) return "RecurringStart changed";
            if (!IsTimeIdentical(GetLocalTime(dstAppointment.End), srcPattern.EndTime)) return "RecurringEnd changed";

            // we create a local temp srcRecurrence to ease up the checks:
            var srcRecurrence = new PatternedRecurrence
            {
                Pattern = new RecurrencePattern(),
                Range = new RecurrenceRange()
            };
            UpdateDestinationPattern(srcPattern, srcRecurrence);

            switch (srcPattern.RecurrenceType)
            {
                case OlRecurrenceType.olRecursWeekly:
                    if (!dstRecurrence.Pattern.DaysOfWeek.SequenceEqual(srcRecurrence.Pattern.DaysOfWeek)) return "Weekly DaysOfWeek changed";
                    break;
                case OlRecurrenceType.olRecursMonthly:
                    if (dstRecurrence.Pattern.DayOfMonth != srcRecurrence.Pattern.DayOfMonth) return "Monthly DayOfMonth changed";
                    break;
                case OlRecurrenceType.olRecursMonthNth:
                    if (dstRecurrence.Pattern.Index != srcRecurrence.Pattern.Index) return "MonthNth Index changed";
                    if (!dstRecurrence.Pattern.DaysOfWeek.SequenceEqual(srcRecurrence.Pattern.DaysOfWeek)) return "MonthlyNth DaysOfWeek changed";
                    break;
                case OlRecurrenceType.olRecursYearly:
                    if (dstRecurrence.Pattern.DayOfMonth != srcRecurrence.Pattern.DayOfMonth) return "Yearly DayOfMonth changed";
                    if (dstRecurrence.Pattern.Month != srcRecurrence.Pattern.Month) return "Yearly Month changed";
                    break;
                case OlRecurrenceType.olRecursYearNth:
                    if (dstRecurrence.Pattern.Index != srcRecurrence.Pattern.Index) return "YearlyNth Index changed";
                    if (!dstRecurrence.Pattern.DaysOfWeek.SequenceEqual(srcRecurrence.Pattern.DaysOfWeek)) return "YearlyNth DaysOfWeek changed";
                    if (dstRecurrence.Pattern.Month != srcRecurrence.Pattern.Month) return "YearlyNth Month changed";
                    break;
            }

            if (dstRecurrence.Range.StartDate != srcRecurrence.Range.StartDate) return "Range StartDate Changed";
            if (dstRecurrence.Pattern.Interval != srcRecurrence.Pattern.Interval) return "Pattern Interval changed";
            if (srcPattern.NoEndDate)
            {
                if (dstRecurrence.Range.Type != RecurrenceRangeType.NoEnd) return "Pattern NoEndDate changed";
            }
            else if (srcPattern.Occurrences >= 0)
            {
                if (dstRecurrence.Range.Type != RecurrenceRangeType.Numbered) return "Range Type changed";
                if (dstRecurrence.Range.NumberOfOccurrences != srcPattern.Occurrences) return "Range NumberOfOccurrences changed";
            }
            else
            {
                if (dstRecurrence.Range.Type != RecurrenceRangeType.EndDate) return "Range Type changed";
                if( !dstRecurrence.Range.EndDate.Equals(srcRecurrence.Range.EndDate)) return "End Date changed";
            }

            var exceptions = srcPattern.Exceptions;
            try
            {
                // if exceptions is a superset of the already synced exceptions, we do not need to enforce a sync, the new ones will be synced correctly.
                // so we count how many we find:
                int numberOfSyncedExceptions = 0;
                foreach (Microsoft.Office.Interop.Outlook.Exception srcException in exceptions)
                {
                    try
                    {
                        var originalDate = srcException.OriginalDate.ToString("O").Substring(0, 10);

                        if (!myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.ContainsKey(originalDate))
                        {
                            continue;
                        }

                        numberOfSyncedExceptions++;
                        if (!srcException.Deleted && myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].Id == null)
                        {
                            return "Series Exception inconsistent sync, forcing a full sync.";
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(srcException);
                    }
                }
                if (numberOfSyncedExceptions != myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.Count)
                {
                    return "Series Exceptions count Changed";
                }
            }
            finally
            {
                Marshal.ReleaseComObject(exceptions);
            }

            return null;
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
                WriteMessageLog("Deleting remote Appointment [{0}].", item.Value.Subject);
                try
                {
                    myMappingTable.Remove(item.Key);
                    await item.Value.DeleteAsync();
                    deletedItems++;
                }
                catch (Exception ex)
                {
                    WriteMessageLog("  ERROR: Could not delete remote appointment [{0}]: {1}", item.Value.Subject, ex.Message);
                }                
            }
            return deletedItems;
        }

        private static WeekIndex GetWeekIndex(int p)
        {
            return (WeekIndex) (p-1);
        }

        private static IList<DayOfWeek> CreateDaysOfWeekList(OlDaysOfWeek olDaysOfWeekMask)
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

        private static RecurrencePatternType GetPatternType(OlRecurrenceType olRecurrenceType)
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

        private static DateTimeTimeZone CreateDateTimeTimeZone(DateTime dateTime)
        {
            return new DateTimeTimeZone
            {
                DateTime = dateTime.ToString("O"),
                TimeZone = TimeZoneInfo.Local.Id
            };
        }

        private static DateTime GetLocalTime(DateTimeTimeZone dateTime)
        {
            var dt = DateTime.Parse(dateTime.DateTime);
            var timeZoneInfo = TimeZoneInfo.Utc;
            if (!string.IsNullOrWhiteSpace(dateTime.TimeZone))
            {
                timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(dateTime.TimeZone);
            }
            return TimeZoneInfo.ConvertTime(dt, timeZoneInfo, TimeZoneInfo.Local);            
        }

        private static bool IsTimeIdentical(DateTime srcDateTime, DateTime dstDateTime)
        {
            if (srcDateTime.Hour != dstDateTime.Hour) return false;
            if (srcDateTime.Minute != dstDateTime.Minute) return false;
            return true;
        }

        private static FreeBusyStatus GetFreeBusyStatus(OlBusyStatus olBusyStatus)
        {
            switch (olBusyStatus)
            {
                case OlBusyStatus.olBusy:
                    return FreeBusyStatus.Busy;
                case OlBusyStatus.olOutOfOffice:
                    return FreeBusyStatus.Oof;
                case OlBusyStatus.olTentative:
                    return FreeBusyStatus.Tentative;
                //case OlBusyStatus.olFree:
                //    return FreeBusyStatus.Free;
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
            if (srcAppointment.IsRecurring)
            {
                var pattern = srcAppointment.GetRecurrencePattern();
                var exceptions = pattern.Exceptions;
                try
                {
                    foreach (Microsoft.Office.Interop.Outlook.Exception exception in exceptions)
                    {
                        try
                        {
                            var originalDate = exception.OriginalDate.ToString("O").Substring(0, 10);
                            if (!myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.ContainsKey(originalDate))
                            {
                                myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.Add(originalDate, new EventIdentificationInfo());
                            }
                            if (!exception.Deleted)
                            {
                                var exceptionItem = exception.AppointmentItem;
                                myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].LastSyncTimeStamp =
                                    exceptionItem.LastModificationTime.ToString("O");
                                Marshal.ReleaseComObject(exceptionItem);
                            }
                            else
                            {
                                myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].LastSyncTimeStamp = null;
                                myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].Id = null;
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(exception);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(exceptions);
                    Marshal.ReleaseComObject(pattern);
                }
            }
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

        private void UpdateExceptionInMappingTable(string srcAppointmentId, string originalDate, string dstAppointmentId)
        {
            if (!myMappingTable.ContainsKey(srcAppointmentId) || myMappingTable[srcAppointmentId].ExceptionIds == null)
            {
                throw new ArgumentException(@"The source appointment must be a recurring, synced appointment", "srcAppointmentId");
            }

            if (!myMappingTable[srcAppointmentId].ExceptionIds.ContainsKey(originalDate))
            {
                myMappingTable[srcAppointmentId].ExceptionIds.Add(originalDate, new EventIdentificationInfo());
            }
            myMappingTable[srcAppointmentId].ExceptionIds[originalDate].Id = dstAppointmentId;
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
            if (String.IsNullOrEmpty(Settings.Default.refresh_token))
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
                    Settings.Default.refresh_token = null;
                    Settings.Default.Save();
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
                    Settings.Default.refresh_token = null;
                    Settings.Default.Save();                        
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
                Settings.Default.refresh_token = null;
                Settings.Default.Save();
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
                    if (cal.Id == Settings.Default.RemoteCaleandarId)
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
            try
            {

                var evts = await myOutlookServicesClient.Me.Calendars[calendarid].Events.ExecuteAsync();
                while (evts != null)
                {
                    result.AddRange(evts.CurrentPage);
                    evts = evts.MorePagesAvailable ? await evts.GetNextPageAsync() : null;
                }
            }
            finally
            {
                RunningAsyncOperations--;
            }

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

            shortcut.TargetPath = Assembly.GetExecutingAssembly().Location;
            shortcut.WorkingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            // shortcut.Arguments = "/Minimized";
            shortcut.WindowStyle = 7;
            shortcut.Description = "Lumis Calendar Sync Autostart";
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

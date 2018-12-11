using IWshRuntimeLibrary;
using LumisCalendarSync.Model;
using LumisCalendarSync.Properties;
using Microsoft.OData.ProxyExtensions;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Threading;
using DayOfWeek = Microsoft.Office365.OutlookServices.DayOfWeek;
using Exception = System.Exception;
using File = System.IO.File;
using RecurrencePattern = Microsoft.Office365.OutlookServices.RecurrencePattern;


namespace LumisCalendarSync.ViewModels
{
    public class MainViewModel: BindableBase
    {
        // This version is shown to the user in the UI and shall be the same as the version set in the MSI.
        // We change this whenever we publish a new MSI version.
        public string CurrentAppVersion
        {
            get
            {
                var assembly = Assembly.GetExecutingAssembly();
                var fileInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
                return fileInfo.FileVersion;
            }
        }
        
        // When we change the list of synced attributes, we change CurrentDataVersion to force a sync of all appointments 
        // at the first sync with the new Application version. 
        // Best Practice: if we change some code (fix or new feature) which needs a full sync, set this to the same value as in the MSI.
        // But not needing to set it for every new MSI version: if no sync must be forced, don't change it.
        private string CurrentDataVersion => "2.15.0.0";


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
                        End = new DateTimeTimeZone {DateTime = (DateTime.Now + TimeSpan.FromHours(1)).ToString("O")},
                        ShowAs = FreeBusyStatus.Tentative
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

        public string LogFileName { get; }

        public ObservableCollection<string> LogEntries { get; }

        private void LogMessage(string format, params object[] arguments)
        {
            var time = DateTime.Now.ToString("s");
            if (String.IsNullOrWhiteSpace(format))
            {
                LogEntries.Add(String.Format(""));
            }
            else
            {
                LogEntries.Add(String.Format("{0}: {1}", time, String.Format(format, arguments)));
            }
            WriteMessageLog(format, arguments);
        }

        private void WriteMessageLog(string format, params object[] arguments)
        {
            // delete the file if it exceeds 1MB
            if (myMessageLog != null && myMessageLog.BaseStream.Length > 1024 * 1024)
            {
                myMessageLog.Close();
                File.Delete(LogFileName);
                myMessageLog = null;
            }
            if (myMessageLog == null)
            {
                try
                {
                    myMessageLog = new StreamWriter(LogFileName, append: true) { AutoFlush = true };
                }
                catch (Exception ex)
                {
                    Error = $"Could not write to log file {LogFileName}, probably another instance of this Application is running";
                    LogEntries.Add(ex.ToString());
                    return;
                }
            }

            string time = DateTime.Now.ToString("s");
            if (String.IsNullOrWhiteSpace(format))
            {
                myMessageLog.WriteLine();
            }
            else
            {
                myMessageLog.WriteLine("{0}: {1}", time, String.Format(format, arguments));
            }
        }

        private void SaveMappingTable()
        {
            var mappingFile = Path.Combine(myAppDataFolder, $"{User.EmailAddress}-{SelectedCalendar.Name}.mapping");
            try
            {
                var serializer = new JavaScriptSerializer();
                File.WriteAllText(mappingFile, serializer.Serialize(myMappingTable));
            }
            catch (Exception ex)
            {
                LogMessage($"Could not save Mapping file {mappingFile}: {ex}");
                Error = "Could not save Mapping File";
            }
        }

        private void LoadMappingTable()
        {
            // Legacy: versions up to 2.0.8 where using a hash code of the calendar id for the file name, which is unstable
            // (different between debug and release, can be changed from .net version)
            // So we move it to a more reliable filename:
            var legacyMappingFile = Path.Combine(myAppDataFolder, $"Mapping-{SelectedCalendar.Id.GetHashCode()}.dat");
            var mappingFile = Path.Combine(myAppDataFolder, $"{User.EmailAddress}-{SelectedCalendar.Name}.mapping");

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
                LogMessage($"Could not load Mapping file {mappingFile}: {ex}");
                Error = "Could not load Mapping File";
            }
        }

        private string myError;
        public string Error
        {
            get => myError;
            set
            {
                Set(ref myError, value, "Error"); 
                if (UserNotification != null && !string.IsNullOrWhiteSpace(value))
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
            get => mySelectedCalendar;
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
            get => mySelectedEvent;
            set
            {
                Set(ref mySelectedEvent, value, "SelectedEvent");
                NotifyCommands();
            }
        }

        public bool IsAutoSyncEnabled
        {
            get => Settings.Default.IsAutoSyncEnabled;
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
            get => Settings.Default.AutoSyncInterval;
            set
            {
                if(value < 10 ) throw new Exception("Provide a value in minutes > 10");
                Settings.Default.AutoSyncInterval = value;
                Settings.Default.Save();
                myTimer.Interval = TimeSpan.FromMinutes(value);
                RaisePropertyChanged("AutoSyncInterval");
            }
        }

        public bool CanAutosync => IsLoggedIn && SelectedCalendar != null;

        public bool CanChangeCalendar => IsLoggedIn && IsIdle && !IsAutoSyncEnabled;

        private void NotifyCommands()
        {
            DeleteAllCommand.RaiseCanExecuteChanged();
            DeleteEventCommand.RaiseCanExecuteChanged();
            SynchronizeCommand.RaiseCanExecuteChanged();
        }

        private IUser myUser;

        public IUser User
        {
            get => myUser;
            set => Set(ref myUser, value, "User");
        }

        private int myRunningAsyncOperations;

        private int RunningAsyncOperations
        {
            get => myRunningAsyncOperations;
            set
            {
                //if (value < 0) value = 0;
                Set(ref myRunningAsyncOperations, value, "RunningAsyncOperations"); 
                RaisePropertyChanged("IsIdle");
                RaisePropertyChanged("CanChangeCalendar");
                NotifyCommands();
            }
        }

        public bool IsIdle => RunningAsyncOperations == 0;

        private bool myIsLoggedIn;
        public bool IsLoggedIn
        {
            get => myIsLoggedIn;
            set
            {
                Set(ref myIsLoggedIn, value, "IsLoggedIn"); 
                RaisePropertyChanged("IsLoggedOut");
                RaisePropertyChanged("CanChangeCalendar");
                RaisePropertyChanged("CanAutosync");
                NotifyCommands();
            }
        }

        public bool IsLoggedOut => !myIsLoggedIn;

        private DelegateCommand myDeleteAllCommand;
        public DelegateCommand DeleteAllCommand
        {
            get
            {
                return myDeleteAllCommand ?? (myDeleteAllCommand = DelegateCommand.FromAsyncHandler(
                                                  executeMethod: async () =>
                                                  {
                                                      await DeleteAllEventsAsync();
                                                  },
                                                  canExecuteMethod: () => SelectedCalendar != null && IsIdle && IsLoggedIn));
            }
        }

        private DelegateCommand myDeleteEventCommand;

        public DelegateCommand DeleteEventCommand
        {
            get
            {
                return myDeleteEventCommand ?? (myDeleteEventCommand = DelegateCommand.FromAsyncHandler(
                                                    executeMethod: () => DeleteEventAsync(SelectedEvent),
                                                    canExecuteMethod: () => SelectedEvent != null && IsIdle && IsLoggedIn));
            }
        }

        private DelegateCommand mySynchronizeCommand;
        public DelegateCommand SynchronizeCommand
        {
            get
            {
                return mySynchronizeCommand ?? (mySynchronizeCommand = DelegateCommand.FromAsyncHandler(
                                                    executeMethod: SynchronizeAsync,
                                                    canExecuteMethod: () => SelectedCalendar != null && IsIdle && IsLoggedIn
                                                ));
            }
        }

        public bool RunAtStartup
        {
            get => IsStartupShortcutSaved();
            set
            {
                if(value) CreateStartupFolderShortcut();
                else DeleteStartupFolderShortcut();
                RaisePropertyChanged("RunAtStartup");
            }
        }

        public bool SkipOldAppointments
        {
            get => Settings.Default.SkipOldAppointments;
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

        private async Task SynchronizeAsync()
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

            RunningAsyncOperations++;
            Items srcAppointmentItems = null;
            try
            {
                using (var outlookWrapper = new OutlookWrapper())
                {
                    srcAppointmentItems = outlookWrapper.GetAppointmentItems();
                    if (srcAppointmentItems == null)
                    {
                        Error = "Outlook is not running, cannot sync.";
                        LogMessage(Error);
                        return;
                    }

                    Events.Clear();
                    LogEntries.Clear();
                    LogMessage($"Syncing your local appointments to remote calendar [{SelectedCalendar.Name}] on account [{User.EmailAddress}].");

                    var remoteCalendarEvents = myOutlookServicesClient.Me.Calendars[Settings.Default.RemoteCaleandarId].Events;
                    var dstAppointmentItems = await GetCalendarEventsAsync(Settings.Default.RemoteCaleandarId);

                    var targetItems = GetAllSyncedEvents(dstAppointmentItems);

                    foreach (var e in dstAppointmentItems)
                    {
                        if (targetItems.Values.Contains(e)) continue;
                        Events.Add(new EventModel(e) { IsSynchronized = false });
                    }

                    // this holds the GlobalAppointmentIDs of those local recurring appointments for which we need so sync exceptions
                    var recurringItemsGlobalIds = new List<string>();

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
                            operationChain += "Checking if target appointment already exists; ";
                            string reasonForSync = "New Appointment";
                            if (String.IsNullOrEmpty(srcAppointment.GlobalAppointmentID))
                            {
                                LogMessage($"    [{currentSubject}] has no Global Appointment ID, ignoring it.");
                                continue;
                            }

                            if (srcAppointment.IsRecurring)
                            {
                                recurringItemsGlobalIds.Add(srcAppointment.GlobalAppointmentID);
                            }

                            if (targetItems.ContainsKey(srcAppointment.GlobalAppointmentID))
                            {
                                dstAppointment = targetItems[srcAppointment.GlobalAppointmentID];
                                targetItems.Remove(srcAppointment.GlobalAppointmentID);

                                // skip appointments which did not changed since last sync
                                if (GetLastSyncTimeStamp(srcAppointment) == srcAppointment.LastModificationTime.ToString("O") && !Settings.Default.ForceNextSync)
                                {
                                    unchangedAppointments++;
                                    Events.Add(new EventModel(dstAppointment){IsSynchronized = true});
                                    continue;
                                }

                                reasonForSync = Settings.Default.ForceNextSync?
                                                    "Application Updated" : DidAppointmentInformationChanged(dstAppointment, srcAppointment);

                                // if nothing important changed, no sync is needed. Just update the last synced time stamp
                                if (reasonForSync == null)
                                {
                                    SetLastSyncTimeStamp(srcAppointment);
                                    unchangedAppointments++;
                                    Events.Add(new EventModel(dstAppointment) { IsSynchronized = true });
                                    continue;
                                }

                                // changed appointments which are Recurring or the IsAllDay attribute changed, are deleted since we cannot update them correctly.
                                // (so we handle them as if they would not have been synced before)
                                if (srcAppointment.IsRecurring || dstAppointment.Recurrence != null || srcAppointment.AllDayEvent != dstAppointment.IsAllDay)
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

                                // Non-recurring appointment:
                                if (!srcAppointment.IsRecurring)
                                {
                                    LogMessage($"    [{srcAppointment.Subject}]: on {srcAppointment.Start}. {reasonForSync}. ");

                                    operationChain += "Not Recurring; ";

                                    operationChain += "Updating Start and End; ";
                                    if (srcAppointment.AllDayEvent)
                                    {
                                        dstAppointment.Start = CreateDateTimeTimeZone(srcAppointment.StartInStartTimeZone.Date, srcAppointment.StartTimeZone.ID);
                                        dstAppointment.End = CreateDateTimeTimeZone(srcAppointment.EndInEndTimeZone.Date, srcAppointment.EndTimeZone.ID);
                                    }
                                    else
                                    {
                                        dstAppointment.Start = CreateDateTimeTimeZone(srcAppointment.StartInStartTimeZone, srcAppointment.StartTimeZone.ID);
                                        dstAppointment.End = CreateDateTimeTimeZone(srcAppointment.EndInEndTimeZone, srcAppointment.EndTimeZone.ID);
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

                                    var dstRecurrence = dstAppointment.Recurrence ?? (dstAppointment.Recurrence = new PatternedRecurrence
                                    {
                                        Pattern = new RecurrencePattern(),
                                        Range = new RecurrenceRange()
                                    });

                                    operationChain += "Updating RecurreneType; ";
                                    dstRecurrence.Pattern.Type = GetPatternType(srcPattern.RecurrenceType);

                                    operationChain += "Updating Start and End; ";
                                    if (srcAppointment.AllDayEvent)
                                    {
                                        dstAppointment.Start = CreateDateTimeTimeZone(srcAppointment.StartInStartTimeZone.Date, srcAppointment.StartTimeZone.ID);
                                        dstAppointment.End = CreateDateTimeTimeZone(srcAppointment.EndInEndTimeZone.Date, srcAppointment.EndTimeZone.ID);
                                    }
                                    else
                                    {
                                        dstAppointment.Start = CreateDateTimeTimeZone(srcAppointment.StartInStartTimeZone, srcAppointment.StartTimeZone.ID);
                                        dstAppointment.End = CreateDateTimeTimeZone(srcAppointment.EndInEndTimeZone, srcAppointment.EndTimeZone.ID);
                                    }

                                    LogMessage($"    [{srcAppointment.Subject}]: recurring {dstRecurrence.Pattern.Type} at {dstAppointment.Start.DateTime.Substring(11, 8)}. {reasonForSync}. ");

                                    UpdateDestinationPattern(srcPattern, dstRecurrence);

                                    operationChain += "Saving; ";

                                    await remoteCalendarEvents.AddEventAsync(dstAppointment);

                                    dstAppointmentItems.Add(dstAppointment);
                                    AddToMappingTable(srcAppointment, dstAppointment);

                                    Events.Add(new EventModel(dstAppointment) { IsSynchronized = true });
                                }
                                successfullyUpdated++;
                                SetLastSyncTimeStamp(srcAppointment);
                            }
                            catch (Exception ex)
                            {
                                if (dstAppointmentIsNew)
                                {
                                    LogMessage($"    ERROR: Could not create appointment [{srcAppointment.Subject}] in target calendar.");
                                    LogMessage($"    Chain of performed operations: {operationChain}.");
                                    LogMessage($"    {ex}");
                                    LogMessage("");
                                }
                                else
                                {
                                    LogMessage($"    ERROR: Could not sync appointment [{srcAppointment.Subject}].");
                                    LogMessage($"    {ex}");
                                    LogMessage("");
                                }
                                errorUpdated++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage($"ERROR syncing [{currentSubject}]. The message below might help us understanding what happened. Sorry.");
                            LogMessage($"Chain of performed operations: {operationChain}.");
                            LogMessage($"{ex}");
                            LogMessage("");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(srcAppointment);
                        }
                    }

                    int deletedItems = await DeleteAppointments(targetItems);

                    await Task.Factory.StartNew(() => Thread.Sleep(1000));

                    await SynchronizeExceptionsAsync(outlookWrapper, recurringItemsGlobalIds);

                    SaveMappingTable();

                    deletedUpdates += deletedItems;
                    errorUpdated += (targetItems.Count - deletedItems);

                    LogMessage("Sync done.");
                    if( successfullyUpdated != 0) LogMessage($"{successfullyUpdated} appointments updated / created.");
                    if (deletedUpdates != 0) LogMessage($"{deletedUpdates} appointments deleted.");
                    if (errorUpdated != 0) LogMessage($"{errorUpdated} appointments failed to be updated.");
                    if (unchangedAppointments != 0) LogMessage($"{unchangedAppointments} appointments did not changed since their last sync.");
                    if(skippedAppointments != 0) LogMessage($"{skippedAppointments} appointments not synced or deleted because they are older than 30 days.");
                    LogMessage("");

                    Settings.Default.ForceNextSync = false;
                    Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error during synchronization: {ex}");
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

        private async Task SynchronizeExceptionsAsync(OutlookWrapper outlookWrapper, List<string> globalIdsToBeSyced)
        {
            Items srcAppointmentItems = null;
            try
            {
                srcAppointmentItems = outlookWrapper.GetAppointmentItems();
                LogMessage("Syncing exceptions (deleted instances or changed instances) for recurring appointments.");

                var remoteCalendarEvents = myOutlookServicesClient.Me.Calendars[Settings.Default.RemoteCaleandarId].Events;
                var dstAppointmentItems = await GetCalendarEventsAsync(Settings.Default.RemoteCaleandarId);
                var targetItems = GetAllSyncedEvents(dstAppointmentItems);

                foreach (var item in srcAppointmentItems)
                {
                    AppointmentItem srcAppointment = null;
                    try
                    {
                        srcAppointment = item as AppointmentItem;
                        if (srcAppointment == null) continue; // finally{} will release the com object 

                        if (!globalIdsToBeSyced.Contains(srcAppointment.GlobalAppointmentID)) continue;

                        if (!targetItems.ContainsKey(srcAppointment.GlobalAppointmentID))
                        {
                            LogMessage($"  ERROR: could not find the remote Appointment for [{srcAppointment.Subject}]. Ignoring it and continue.");
                            continue;
                        }

                        var dstAppointment = targetItems[srcAppointment.GlobalAppointmentID];
                        targetItems.Remove(srcAppointment.GlobalAppointmentID);

                        var srcPattern = srcAppointment.GetRecurrencePattern();

                        // Handling exceptions for the recurring appointment:
                        var srcExceptions = GetRelevantExceptions(srcPattern.Exceptions);
                        var numberOfUnchangedExceptions = 0;
                        var numberOfChangedExceptions = 0;
                        foreach (var srcException in srcExceptions)
                        {
                            var srcExceptionItem = srcException.Deleted ? null : srcException.AppointmentItem;
                            try
                            {
                                IEvent dstExceptionItem = null;
                                var originalDate = srcException.OriginalDate.ToString("O").Substring(0, 10);
                                if (myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.ContainsKey(originalDate))
                                {
                                    bool isDeletedAndSynced = srcException.Deleted 
                                                              && myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].Id 
                                                              == null;
                                    bool isOtherAndSynced = !srcException.Deleted
                                                            && myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].LastSyncTimeStamp
                                                            == srcExceptionItem.LastModificationTime.ToString("O");

                                    if (isDeletedAndSynced || isOtherAndSynced)
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

                                numberOfChangedExceptions++;
                                if(dstExceptionItem == null)
                                {
                                    var intervalStart = srcException.OriginalDate - TimeSpan.FromDays(1);
                                    var intervalEnd = srcException.OriginalDate + TimeSpan.FromDays(1);
                                    var eventCollection = await remoteCalendarEvents[dstAppointment.Id].GetInstances(
                                                              new DateTimeOffset(intervalStart), new DateTimeOffset(intervalEnd)).ExecuteAsync();
                                    var remoteInstances = await GetEventInstancesAsync(eventCollection);
                                    dstExceptionItem = remoteInstances.FirstOrDefault(ri =>
                                        GetLocalTime(ri.Start).ToString("O").Substring(0, 10) == originalDate);
                                }

                                if (dstExceptionItem == null)
                                {
                                    LogMessage($"  [{srcAppointment.Subject}]: ERROR. No remote instance found for local {(srcException.Deleted ? "deleted " : "")}exception on {originalDate}.");
                                    continue;
                                }

                                if (srcException.Deleted)
                                {
                                    await dstExceptionItem.DeleteAsync();
                                    UpdateExceptionInMappingTable(srcAppointment.GlobalAppointmentID, originalDate, null);
                                }
                                else
                                {
                                    dstExceptionItem.SeriesMasterId = dstAppointment.Id;
                                    dstExceptionItem.Type = EventType.Exception;
                                    dstExceptionItem.Subject = srcExceptionItem.Subject;
                                    dstExceptionItem.Location = new Location {DisplayName = srcExceptionItem.Location};
                                    dstExceptionItem.Start = CreateDateTimeTimeZone(srcExceptionItem.StartInStartTimeZone, srcExceptionItem.StartTimeZone.ID);
                                    dstExceptionItem.End = CreateDateTimeTimeZone(srcExceptionItem.EndInEndTimeZone, srcExceptionItem.EndTimeZone.ID);
                                    dstExceptionItem.IsReminderOn = srcExceptionItem.ReminderSet;
                                    if (srcExceptionItem.ReminderSet)
                                    {
                                        dstExceptionItem.ReminderMinutesBeforeStart = srcExceptionItem.ReminderMinutesBeforeStart;
                                    }

                                    await dstExceptionItem.UpdateAsync();
                                    var lastChange = srcExceptionItem.LastModificationTime.ToString("O");

                                    UpdateExceptionInMappingTable(srcAppointment.GlobalAppointmentID, originalDate, dstExceptionItem.Id, lastChange);
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

                        if (numberOfChangedExceptions > 0)
                        {
                            LogMessage($"    [{srcAppointment.Subject}]: {numberOfChangedExceptions} exceptions have been synced, {numberOfUnchangedExceptions} unchanged since last sync.");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"  ERROR: Could not sync appointment [{srcAppointment?.Subject}].");
                        LogMessage($"  {ex}");
                        LogMessage("");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error during synchronization: {ex}");
                Error = "Error during synchronization, see log file.";
            }
            finally
            {
                if (srcAppointmentItems != null)
                {
                    Marshal.ReleaseComObject(srcAppointmentItems);
                }
            }
        }

        private List<Microsoft.Office.Interop.Outlook.Exception> GetRelevantExceptions(Exceptions srcExceptions)
        {
            var result = new List<Microsoft.Office.Interop.Outlook.Exception>();

            try
            {
                foreach (Microsoft.Office.Interop.Outlook.Exception srcException in srcExceptions)
                {
                    if (SkipOldAppointments)
                    {
                        if (srcException.Deleted && (DateTime.Now.Date - srcException.OriginalDate.Date).TotalDays > 30)
                        {
                            Marshal.ReleaseComObject(srcException);
                            continue;
                        }

                        var srcExceptionItem = srcException.Deleted ? null : srcException.AppointmentItem;
                        if (srcExceptionItem != null)
                        {
                            var endDate = srcExceptionItem.End.Date;
                            Marshal.ReleaseComObject(srcExceptionItem);
                            if ((DateTime.Now.Date - endDate).TotalDays > 30)
                            {
                                Marshal.ReleaseComObject(srcException);
                                continue;
                            }
                        }

                        result.Add(srcException);
                    }
                }
                return result;
            }
            finally
            {
                Marshal.ReleaseComObject(srcExceptions);
            }
        }

        private static bool IsAppointmentOld(AppointmentItem srcAppointment)
        {
            bool isToBeSkipped;
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
                    // Example: every second Tuesday of the month, every 2 months would be:
                    // Index will be 1 (first == 0, second == 1, ...)
                    // DaysOfWeek will be Tuesday
                    // Interval will be 2 (every two months), covered below the switch for all cases.
                    dstRecurrence.Pattern.Index = GetWeekIndex(srcPattern.Instance);
                    dstRecurrence.Pattern.DaysOfWeek = CreateDaysOfWeekList(srcPattern.DayOfWeekMask);
                    break;
                case OlRecurrenceType.olRecursYearly:
                    dstRecurrence.Pattern.DayOfMonth = srcPattern.DayOfMonth;
                    dstRecurrence.Pattern.Month = srcPattern.MonthOfYear;
                    break;
                case OlRecurrenceType.olRecursYearNth:
                    // example: every 2nd Tuesday of January, every 2 years would be:
                    // index = 1 (0 is first, 1 is second, ...),
                    // DaysOfWeek = Tuesday 
                    // Month would be 1 (January is 1, December is 12)
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

        private string DidAppointmentInformationChanged(IEvent dstAppointment, AppointmentItem srcAppointment)
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
            if (!IsTimeIdentical(GetLocalTime(dstAppointment.Start), srcAppointment.Start)) return "RecurringStart changed";
            if (!IsTimeIdentical(GetLocalTime(dstAppointment.End), srcAppointment.End)) return "RecurringEnd changed";

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
                // We need to force a full sync of a recurring appointment when we have deleted in the past an exception and now it is the "kind" of an exception changes.
                // New exceptions and attribute change of exceptions will be handled without a forced sync...
                int numberOfSyncedExceptions = 0;
                foreach (var srcException in GetRelevantExceptions(exceptions))
                {
                    try
                    {
                        var originalDate = srcException.OriginalDate.ToString("O").Substring(0, 10);

                        if (!myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.ContainsKey(originalDate))
                        {
                            continue; // new exception, will be synced automatically
                        }
                        numberOfSyncedExceptions++;

                        if (!srcException.Deleted && myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds[originalDate].Id == null)
                        {
                            return "Series Exception changed";
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(srcException);
                    }
                }
                if (numberOfSyncedExceptions != myMappingTable[srcAppointment.GlobalAppointmentID].ExceptionIds.Count)
                {
                    return "Series Exceptions changed";
                }
            }
            finally
            {
                Marshal.ReleaseComObject(exceptions);
            }

            return null;
        }

        private async Task<List<IEvent>> GetEventInstancesAsync(IPagedCollection<IEvent> eventCollection)
        {
            var result = new List<IEvent>();
            while (eventCollection != null)
            {
                result.AddRange(eventCollection.CurrentPage);
                eventCollection = eventCollection.MorePagesAvailable ? await eventCollection.GetNextPageAsync() : null;
            }
            return result;
        }

        private async Task<int> DeleteAppointments(Dictionary<string, IEvent> targetItems)
        {
            var deletedItems = 0;
            foreach (var item in targetItems)
            {
                LogMessage($"    Deleting remote Appointment [{item.Value.Subject}].");
                try
                {
                    myMappingTable.Remove(item.Key);
                    await item.Value.DeleteAsync();
                    deletedItems++;
                }
                catch (Exception ex)
                {
                    LogMessage($"    ERROR: Could not delete remote appointment [{item.Value.Subject}]: {ex.Message}");
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

        private static DateTimeTimeZone CreateDateTimeTimeZone(DateTime dateTime, string timeZoneId)
        {
            return new DateTimeTimeZone
            {
                DateTime = dateTime.ToString("O"),
                TimeZone = timeZoneId
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
                case OlBusyStatus.olFree:
                    return FreeBusyStatus.Free;
                default:
                    return FreeBusyStatus.Unknown;
            }
        }

        private string GetLastSyncTimeStamp(AppointmentItem srcAppointment)
        {
            //return dstAppointment.UserProperties["originalLastUpdate"];
            if (!myMappingTable.ContainsKey(srcAppointment.GlobalAppointmentID)) return null;
            return myMappingTable[srcAppointment.GlobalAppointmentID]?.LastSyncTimeStamp;
        }


        private void SetLastSyncTimeStamp(AppointmentItem srcAppointment)
        {
            myMappingTable[srcAppointment.GlobalAppointmentID].LastSyncTimeStamp = srcAppointment.LastModificationTime.ToString("O");
        }

        private void UpdateExceptionInMappingTable(string srcAppointmentId, string originalDate, string dstAppointmentId, string lastChange = null)
        {
            if (!myMappingTable.ContainsKey(srcAppointmentId))
            {
                throw new ArgumentException(@"The source appointment is not found in the must have been already synced", nameof(srcAppointmentId));
            }

            if (!myMappingTable[srcAppointmentId].ExceptionIds.ContainsKey(originalDate))
            {
                myMappingTable[srcAppointmentId].ExceptionIds.Add(originalDate, new EventIdentificationInfo());
            }
            myMappingTable[srcAppointmentId].ExceptionIds[originalDate].Id = dstAppointmentId;
            myMappingTable[srcAppointmentId].ExceptionIds[originalDate].LastSyncTimeStamp = lastChange;
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

        private async Task DeleteEventAsync(EventModel e)
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
        private async Task DeleteAllEventsAsync()
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
                LogMessage($"Could not delete all events: {exception}");
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

        private async void AutologinAsync()
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
                LogMessage($"Could not log you in: {ex}");
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

        public async void LoginAsync(string authCode)
        {
            if (String.IsNullOrEmpty(authCode))
            {
                throw new ArgumentNullException(nameof(authCode));
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
                LogMessage($"Could not log you in: {ex}");
                Error = "Login Failed";
            }
            finally
            {
                RunningAsyncOperations--;
            }
        }

        private void LogApiMigrationError()
        {
            LogMessage("");
            LogMessage("Lumis Calendar sync is using a new Microsoft API, and not all Microsoft accounts are enabled for this API.");
            LogMessage("The most probable cause for the error above is that your account has not yet been enabled.");
            LogMessage("It is not clear how log it will take until your account will be migrated.");
            LogMessage(
                "Currently, the only workaround is for you to create a new Microsoft account and use it, as all new accounts have the new API automatically enabled.");
            LogMessage(
                "Alternatively, continue to use the previous version of Lumis Calendar Sync and try this new version periodically until your account gets migrated.");
            Error = "We cannot access your Outlook.com Calendars. Your account might not have been migrated yet.";
        }

        public async void LogoutAsync()
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
                LogMessage($"Error logging you out: {exception}");
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

        private async Task GetUserDataAsync()
        {
            User = await myOutlookServicesClient.Me.ExecuteAsync();
        }

        private async Task GetUserCalendarsAsync()
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

        private async void PopulateEventsAsync(string calendarId)
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
                LogMessage($"Error in retrieving the remote events: {ex}");
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

        private async Task<IList<IEvent>> GetCalendarEventsAsync(string calendarid)
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
        private readonly string myAppDataFolder;
    }
}

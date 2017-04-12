**Project Description**

Syncs your Outlook appointments from work to your private calendar on outlook.com, so you can get reminders on your smartphone.

If you use outlook desktop app at work (with an exchange server), but cannot sync your appointments with your smartphone (e.g. because your company is not allowing your phone to sync with the companies exchange server), Lumis Calendar Sync might help you without infringing the companies policy:
It syncs only a minimal amount of information **from** the appointments in your default local outlook calendar **to** one of your (private) calendars on outlook.com.

Starting with Version 2.0, Lumis Calendar Sync is using a new API (called the Office 365 API) to access your cloud calendar.

**Prerequisite:**

your Microsoft account (containing your target calendar) must have been migrated to the office 365 servers. [Read more about the new Outlook.com](https://blogs.office.com/2016/02/17/outlook-com-out-of-preview-and-better-than-ever/).
Unfortunately, you cannot speed up or influence when your account will be migrated, currently known target is 31. August 2016.
You can verify if your account has been migrated by logging in to your account at [http://outlook.com](http://outlook.com) and verify if the label above your inbox is "Outlook-Mail". If yes, your account runs on the office 365 servers. If it is "Outlook.com", you are on the old outlook.com servers and this version of Lumis Calendar Sync will not work for you. 

**Notes:** 

if you create a new Microsoft account, it is hosted on the new office 365 servers!

If your Microsoft account has not been migrated yet to the Office 365 servers, and you do not want to create a new Microsoft account, you can (continue to) use Lumis Calendar Sync version 1.4 from [http://LumisCalendarSync.CodePlex.com](CodePlex): select manually the version 1.4 on the download page and follow this guide: [Lumis-Calendar-Sync-up-to-and-including-Version-1.4-(running-as-Outlook-AddIn,-requiering-OutlookConnector)](https://lumiscalendarsync.codeplex.com/wikipage?title=Lumis%20Calendar%20Sync%20up%20to%20and%20including%20Version%201.4%20%28running%20as%20Outlook%20AddIn%2c%20requiering%20OutlookConnector%29&referringTitle=Home)

**How it works (Version 2.0 and higher)**

Lumis Calendar Sync runs as a regular windows program, which can be minimized to the system tray, and it is accessing your running outlook desktop application on the one side and your outlook.com calendar on the other side, in order to sync the subject and meeting room from your appointments (no participants, no content) and only in one direction (from your default local outlook calendar to your private outlook.com calendar) - just enough to provide you with reminders on your smartphone.

But is a real sync: it is able to keep track of the changes in your local calendar (creation, updates, deletions) and propagate them periodically to your private calendar. It is able to deal with recurring appointments, including exceptions (like deletion of a particular instance of a series or move of a particular instance of a series to another time or another meeting room), so that your synced calendar is always up to date.

**How to use it**

* I recommend to use a dedicated calendar for this purpose - it will avoid mixing up your private and work appointments in one calendar: Go to [outlook.com](http://outlook.com), log in with your Microsoft account, go to Calendar and create a new calendar (name it e.g. "Calendar@Work") where you want to sync your work appointments. 
* Install Lumis Calendar Sync, login in with your Microsoft account, select your target calendar (the one you have created in the previous step) and sync. You can configure Lumis Calendar Sync to start when you log in and to periodically sync your appointments - as long as your computer is not sleeping, shutting down or logging you out.

**Note:** make sure your outlook desktop application is running.

Connecting your smartphone to your calendar on outlook.com will provide you with reminders on your smartphone.

**Upgrade Notes:** 

If you where using Lumis Calendar Sync 1.4 or earlier, please:
* uninstall it before using this version. 
* remove your outlook.com calendar which you configured with outlook connector in your outlook desktop application, and uninstall outlook connector - unless you need them for other purposes.
* after logging in in Lumis Calendar Sync and selecting your target calendar, remove all Appointments (events), because thy are not connected to your work appointments and a sync would result in duplicated entries. The first sync will re-create them, including the "connection" information (you can distinguish between the connected and not connected appointments in the table of events and selectively delete some of them at any time later).

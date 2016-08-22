using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;

using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;


namespace LumisCalendarSync
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();

            myTimer = new DispatcherTimer();
            myTimer.Tick += Timer_Tick;
            myTimer.Interval = TimeSpan.FromSeconds(1);

            myNotifyIcon = new NotifyIcon { Icon = Properties.Resources.LumisCalendarSyncIcon, Visible = true };
            myNotifyIcon.DoubleClick += NotifyIcon_OpenClicked;
            myNotifyIcon.ContextMenu = new ContextMenu();
            myNotifyIcon.ContextMenu.MenuItems.Add(new MenuItem { Index = 0, Text = Properties.Resources.OpenMenuCaption });
            myNotifyIcon.ContextMenu.MenuItems[0].Click += NotifyIcon_OpenClicked;
            myNotifyIcon.ContextMenu.MenuItems.Add("-");
            myNotifyIcon.ContextMenu.MenuItems.Add(new MenuItem { Index = 1, Text = Properties.Resources.ShowErrorsMenuCaption });
            myNotifyIcon.ContextMenu.MenuItems[2].Click += NotifyIcon_ShowErrorsClicked;
            myNotifyIcon.ContextMenu.MenuItems[2].Checked = Properties.Settings.Default.ShowErrorNotifications;
            myNotifyIcon.ContextMenu.MenuItems.Add(new MenuItem { Index = 2, Text = Properties.Resources.ShowInfosMenuCaption });
            myNotifyIcon.ContextMenu.MenuItems[3].Click += NotifyIcon_ShowInfosClicked;
            myNotifyIcon.ContextMenu.MenuItems[3].Checked = Properties.Settings.Default.ShowInfoNotifications; 
            myNotifyIcon.ContextMenu.MenuItems.Add("-");
            myNotifyIcon.ContextMenu.MenuItems.Add(new MenuItem { Index = 3, Text = Properties.Resources.ExitMenuCaption });
            myNotifyIcon.ContextMenu.MenuItems[5].Click += NotifyIcon_ExitClicked;

            Title = myNotifyIcon.Text = Properties.Resources.ApplicationCaption;

            MainViewModel.LogEntries.CollectionChanged += LogEntries_CollectionChanged;
            MainViewModel.UserNotification += MainViewModel_UserNotification;
        }

        void NotifyIcon_ShowInfosClicked(object sender, EventArgs e)
        {
            myNotifyIcon.ContextMenu.MenuItems[3].Checked = !myNotifyIcon.ContextMenu.MenuItems[3].Checked;
            Properties.Settings.Default.ShowInfoNotifications = myNotifyIcon.ContextMenu.MenuItems[3].Checked;
            Properties.Settings.Default.Save();
        }

        void NotifyIcon_ShowErrorsClicked(object sender, EventArgs e)
        {
            myNotifyIcon.ContextMenu.MenuItems[2].Checked = !myNotifyIcon.ContextMenu.MenuItems[2].Checked;
            Properties.Settings.Default.ShowErrorNotifications = myNotifyIcon.ContextMenu.MenuItems[2].Checked;
            Properties.Settings.Default.Save();
        }

        private void NotifyIcon_OpenClicked(object sender, EventArgs e)
        {
            Show();
            this.Dispatcher.BeginInvoke(new Action(() => WindowState = WindowState.Normal));
        }

        void NotifyIcon_ExitClicked(object sender, EventArgs e)
        {
            myExitActivated = true;
            Close();
        }

        private bool myExitActivated;
        private readonly NotifyIcon myNotifyIcon;

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == System.Windows.WindowState.Minimized)
            {
                this.Hide();
                if (Properties.Settings.Default.ShowInfoNotifications)
                {
                    myNotifyIcon.ShowBalloonTip(5000, "Lumis Calendar Sync", Properties.Resources.StillRunningText, ToolTipIcon.Info);
                }
            }
            base.OnStateChanged(e);
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (Properties.Settings.Default.IsAutoSyncEnabled && !myExitActivated)
            {
                e.Cancel = true;
                this.WindowState = WindowState.Minimized;
            }
            base.OnClosing(e);
            if (!e.Cancel)
            {
                MainViewModel.Cleanup();
                MainViewModel.Close();
                myNotifyIcon.Dispose();
                Application.Current.Shutdown();
            }
        }
        private void BtnSignIn_Click(object sender, RoutedEventArgs e)
        {
            MainViewModel.Error = "";
            var loginBrowser = new BrowserWindow(MainViewModel.OAuthHelper.LogInUrl);
            loginBrowser.Closed += LoginBrowser_Closed;
            loginBrowser.ShowDialog();
        }

        private void BtnSignOut_Click(object sender, RoutedEventArgs e)
        {
            MainViewModel.LogoutAsync();
        }

        void LoginBrowser_Closed(object sender, EventArgs e)
        {
            var browser = sender as BrowserWindow;
            if (browser != null)
            {
                if (!String.IsNullOrEmpty(browser.ErrorCode))
                {
                    MainViewModel.Error = String.Format("{0}: {1}", browser.ErrorCode, browser.ErrorDescription);
                }
                else if (!String.IsNullOrEmpty(browser.AuthorizationCode))
                {
                    MainViewModel.LoginAsync(browser.AuthorizationCode);
                }
                else
                {
                    MainViewModel.Error = "Log in was canceled";
                }
            }
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
        }

        void MainViewModel_UserNotification(object sender, ViewModels.NotificationEventArgs e)
        {
            if (WindowState == WindowState.Minimized && Properties.Settings.Default.ShowInfoNotifications && !String.IsNullOrEmpty(e.Text))
            {
                myNotifyIcon.ShowBalloonTip(5000, "Lumis Calendar Sync", e.Text, ToolTipIcon.Error);
            }
        }

        private readonly DispatcherTimer myTimer;

        void LogEntries_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            myTimer.Stop();
            myTimer.Start();
        }

        void Timer_Tick(object sender, System.EventArgs e)
        {
            myTimer.Stop();
            if(MainViewModel.LogEntries.Count > 0)
            {
                LogListBox.ScrollIntoView(MainViewModel.LogEntries.Last());
            }
        }

    }
}

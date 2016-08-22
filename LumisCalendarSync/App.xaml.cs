using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;

namespace LumisCalendarSync
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void App_OnStartup(object sender, StartupEventArgs e)
        {
            var startMinimized = false;
            for (var i = 0; i != e.Args.Length; ++i)
            {
                if (e.Args[i] == "/Minimized")
                {
                    startMinimized = true;
                }
            }

            var mainWindow = new MainWindow();
            if (startMinimized)
            {
                mainWindow.WindowState = WindowState.Minimized;
            }
            else
            {
                mainWindow.Show();
            }
        }
    }
}

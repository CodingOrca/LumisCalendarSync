using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LumisCalendarSync
{
    public partial class ConfigurationForm : Form
    {
        ThisAddIn outlookAddIn;
        public ConfigurationForm(ThisAddIn addIn)
        {
            outlookAddIn = addIn;
            InitializeComponent();
            this.Load += new EventHandler(ConfigurationForm_Load);
        }

        void ConfigurationForm_Load(object sender, EventArgs e)
        {
            var calendars = outlookAddIn.GetPossibleDestinationCalendars();
            string selectedItem = LumisCalendarSync.Properties.Settings.Default.DestinationCalendar;
            if( calendars.Count > 0 )
            {
                foreach (string s in calendars)
                {
                    comboBox.Items.Add(s);
                    if (s == selectedItem)
                    {
                        comboBox.SelectedItem = s;
                    }
                }
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            LumisCalendarSync.Properties.Settings.Default.DestinationCalendar = comboBox.Text;
            LumisCalendarSync.Properties.Settings.Default.Save();
            this.Close();
            outlookAddIn.SyncNow();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

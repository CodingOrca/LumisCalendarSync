using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Text.RegularExpressions;
using System.Web;

namespace LumisCalendarSync
{
    public partial class BrowserWindow : Window
    {
        public string AuthorizationCode { get; private set; }
        public string ErrorCode { get; private set; }
        public string ErrorDescription { get; private set; }

        public BrowserWindow(Uri uri)
        {
            InitializeComponent();
            this.AuthorizationCode = null;
            webBrowser.Navigate(uri);
        }

        private void WebBrowser_LoadCompleted(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            // example: when user declines the authorization page (presses no when he is informed what this apps requests to be allowed to do),
            // this .Query part of the uri is returned:
            //
            // ?error=access_denied&error_description=The%20user%20has%20denied%20access%20to%20the%20scope%20requested%20by%20the%20client%20application.&lc=1031
            //
            var queryPairs = HttpUtility.ParseQueryString(e.Uri.Query);
            this.ErrorCode = queryPairs.Get("error");
            if( this.ErrorCode != null)
            {
                this.ErrorDescription = queryPairs.Get("error_description");
                this.Close();
            }

            this.AuthorizationCode = queryPairs.Get("code");
            if (this.AuthorizationCode != null)
            {
                this.Close();
            }
        }
    }
}
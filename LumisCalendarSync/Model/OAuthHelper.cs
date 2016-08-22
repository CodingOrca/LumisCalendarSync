using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace LumisCalendarSync.Model
{
    public partial class OAuthHelper
    {
        public async Task AuthorizeTaskAsync(string authorizationCode)
        {
            using (var client = new HttpClient())
            {
                var values = new Dictionary<string, string>
                {
                    {"grant_type", "authorization_code"},
                    {"code", authorizationCode},
                    {"scope", Scopes},
                    {"redirect_uri", RedirectUrl},
                    {"client_id", ClientID}
                };

                var content = new FormUrlEncodedContent(values);
                var response = await client.PostAsync(TokenUrl, content);
                var result = await response.Content.ReadAsStringAsync();
                ProcessResults(result);
            }
        }

        async public Task<string> GetAccessTokenTaskAsync()
        {
            if (String.IsNullOrEmpty(Properties.Settings.Default.refresh_token)) return null;
            using (var client = new HttpClient())
            {
                var values = new Dictionary<string, string>
                {
                    {"client_id", ClientID},
                    {"redirect_uri", RedirectUrl},
                    {"grant_type", "refresh_token"},
                    {"scope", Scopes},
                    {"refresh_token", Properties.Settings.Default.refresh_token}
                };

                var content = new FormUrlEncodedContent(values);
                var response = await client.PostAsync(TokenUrl, content);
                var result = await response.Content.ReadAsStringAsync();
                return ProcessResults(result);
            }
        }

        async public Task LogoutTaskAsync()
        {
            using (var client = new HttpClient())
            {
                await client.GetStringAsync(myLogoutUrl);
            }
        }

        private string ProcessResults(string downloadedString)
        {
            var tokenData = DeserializeJson(downloadedString);

            if (tokenData.ContainsKey("error"))
            {
                throw new Exception(String.Format("Error {0}: {1}", tokenData["error"], tokenData["error_description"]));
            }
            if (tokenData.ContainsKey("refresh_token"))
            {
                RefreshToken = tokenData["refresh_token"] as string;
                Properties.Settings.Default.refresh_token = RefreshToken;
                Properties.Settings.Default.Save();
            }
            if (tokenData.ContainsKey("access_token"))
            {
                return tokenData["access_token"] as string;
            }
            return null;
        }

        private static Dictionary<string, object> DeserializeJson(string json)
        {
            try
            {
                var jss = new JavaScriptSerializer();
                var d = jss.Deserialize<Dictionary<string, object>>(json);
                return d;
            }
            catch (Exception)
            {
                throw new Exception(json);
            }
        }

        private string RefreshToken { get; set; }

        public readonly Uri LogInUrl =
            new Uri(
                String.Format(@"{0}?client_id={1}&scope={2}&response_type=code&response_mode=query&prompt=login&redirect_uri={3}", AuthorizeUrl, ClientID, UrlEncodedScopes, RedirectUrl));

        private static readonly string myLogoutUrl =
            String.Format(@"{0}?client_id={1}&redirect_uri={2}", LogoutUrl, ClientID, RedirectUrl);
        
        private const string UrlEncodedScopes = "offline_access%20https%3A%2F%2Foutlook.office.com%2Fcalendars.readwrite";
        private const string Scopes = "offline_access https://outlook.office.com/calendars.readwrite";

        private const string RedirectUrl = "urn:ietf:wg:oauth:2.0:oob";
        private const string AuthorizeUrl = "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize";
        private const string TokenUrl = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token";
        private const string LogoutUrl = "https://login.microsoftonline.com/consumers/oauth2/logout?post_logout_redirect_uri=urn:ietf:wg:oauth:2.0:oob";
    }
}

//        //
//        // https://apps.dev.microsoft.com
//        // Make sure you use the ClientId of an application in the section "My applications" at the beginning, NOT a pure Live SDK-Application!
//        // Your application must have a Mobile Application profile, and you should see the Redirect-URI "urn:ietf:wg:oauth:2.0:oob" in its properties.
//        // No secret, no passoword necessary.

// The file OAuthHelper.ClientId.cs is missing from codeplex for security reasons (I do not want to make my ClientId public)
// It must contain the following code:

//namespace LumisCalendarSync.Model
//{
//    public partial class OAuthHelper
//    {
//        private const string ClientID = "PUT_YOUR_CLIENT_ID_HERE";
//    }
//}
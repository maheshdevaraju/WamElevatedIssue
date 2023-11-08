using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.NativeInterop;

namespace WamElevatedIssue
{
    public partial class Ribbon1
    {
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        MyIdentityLogger myLogger = new MyIdentityLogger();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            await LogIn();
        }

        private async Task LogIn()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            var hWnd = (IntPtr)Globals.ThisAddIn.Application.Hwnd;
            var scopes = new[] { "User.Read" };

            IPublicClientApplication app =
                PublicClientApplicationBuilder.Create("Client-Id")
                    .WithDefaultRedirectUri()
                    .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows))
                    .WithLogging(myLogger, true)
                    .Build();
            AuthenticationResult result = null;

            // Try to use the previously signed-in account from the cache
            IEnumerable<IAccount> accounts = await app.GetAccountsAsync();
            IAccount existingAccount = accounts.FirstOrDefault();

            try
            {
                if (existingAccount != null)
                {
                    result = await app.AcquireTokenSilent(scopes, existingAccount).ExecuteAsync();
                }
                // Next, try to sign in silently with the account that the user is signed into Windows
                else
                {
                    result = await app.AcquireTokenSilent(scopes, PublicClientApplication.OperatingSystemAccount)
                        .ExecuteAsync();
                }
            }
            // Can't get a token silently, go interactive
            catch (MsalUiRequiredException ex)
            {
                try
                {
                    result = await app.AcquireTokenInteractive(scopes)
                        .WithParentActivityOrWindow(hWnd)
                        .ExecuteAsync();
                }
                catch (MsalClientException exception)
                {
                    log.Error(exception);
                    //Console.WriteLine(exception);
                }
            }
        }
    }
}

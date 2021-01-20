using System;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

namespace EWSOAuthDelegated
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void WriteToResults(string data)
        {
            // Add the given data to the results textbox

            Action action = new Action(() => {
                textBoxResults.AppendText($"{data}{Environment.NewLine}");
            });
            if (textBoxResults.InvokeRequired)
                textBoxResults.Invoke(action);
            else
                action();
        }

        private void buttonFindFolders_Click(object sender, EventArgs e)
        {
            Action action = new Action(async () =>
            {
                // Configure the MSAL client to get tokens
                var pcaOptions = new PublicClientApplicationOptions
                {
                    ClientId = textBoxAppId.Text,
                    TenantId = textBoxTenantId.Text
                };

                var pca = PublicClientApplicationBuilder
                    .CreateWithApplicationOptions(pcaOptions).Build();

                var ewsScopes = new string[] { "https://outlook.office.com/EWS.AccessAsUser.All" };

                try
                {
                    // Make the interactive token request
                    var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();

                    // Configure the ExchangeService with the access token
                    var ewsClient = new ExchangeService();
                    ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                    ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);

                    // Make an EWS call
                    var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                    foreach (var folder in folders)
                    {
                        WriteToResults($"Folder: {folder.DisplayName}");
                    }
                }
                catch (MsalException ex)
                {
                    WriteToResults($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    WriteToResults($"Error: {ex}");
                }

            });
            System.Threading.Tasks.Task.Run(action);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonGetInboxMessages_Click(object sender, EventArgs e)
        {
            Action action = new Action(async () =>
            {
                // Configure the MSAL client to get tokens
                var pcaOptions = new PublicClientApplicationOptions
                {
                    ClientId = textBoxAppId.Text,
                    TenantId = textBoxTenantId.Text
                };

                var pca = PublicClientApplicationBuilder
                    .CreateWithApplicationOptions(pcaOptions).Build();

                var ewsScopes = new string[] { "https://outlook.office.com/EWS.AccessAsUser.All" };

                try
                {
                    // Make the interactive token request
                    var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();

                    // Configure the ExchangeService with the access token
                    var ewsClient = new ExchangeService(ExchangeVersion.Exchange2016);
                    ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                    ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);

                    // Make an EWS call
                    var items = ewsClient.FindItems(WellKnownFolderName.Inbox,new ItemView(10));
                    foreach (var item in items)
                    {
                        WriteToResults($"{item.DateTimeReceived}: {item.Subject}");
                    }
                }
                catch (MsalException ex)
                {
                    WriteToResults($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    WriteToResults($"Error: {ex}");
                }

            });
            System.Threading.Tasks.Task.Run(action);
        }
    }
}

using System;
using System.Configuration;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

namespace SplunkMailProcessor
{
    class Program
    {
        static void Main(string[] args)
        {   
            var splunkApiKey = ConfigurationManager.AppSettings["Splunk.ApiKey"];
            var splunkCollectorUrl = ConfigurationManager.AppSettings["Splunk.Url"];
            var exchangeUser = ConfigurationManager.AppSettings["Exchange.User"];
            var appId = ConfigurationManager.AppSettings["Exchange.AppId"];
            var clientSecret = ConfigurationManager.AppSettings["Exchange.ClientSecret"];
            var tenantId = ConfigurationManager.AppSettings["Exchange.TenantId"];
            var exchangeUrl = ConfigurationManager.AppSettings["Exchange.Url"];
            var exchangeMsgLimit = int.Parse(ConfigurationManager.AppSettings["Exchange.MessagesPerPoll"]);

            Console.WriteLine("Starting Splunk Mail Processor");
            Console.WriteLine("  Exchange Account: {0}", exchangeUser);
            Console.WriteLine("  Exchange Url: {0}", exchangeUrl);
            Console.WriteLine("  Splunk HTTP Event Collector Url: {0}", splunkCollectorUrl);

            // See these pages for setup within Office365/Azure/Exchange and explaination of auth code
            // https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth
            // https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access

            // Using Microsoft.Identity.Client 
            var cca = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithClientSecret(clientSecret)
                .WithTenantId(tenantId)
                .Build();

            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

            // Make the token request
            var authResult = cca.AcquireTokenForClient(ewsScopes).ExecuteAsync().Result;

            // Use token to authenticate with exchange service
            var exchangeService = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Credentials = new OAuthCredentials(authResult.AccessToken),
                Url = new Uri(exchangeUrl)
            };

            //Impersonate the mailbox to access.
            exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, exchangeUser);

            // Include x-anchormailbox header
            exchangeService.HttpHeaders.Add("X-AnchorMailbox", exchangeUser);

            var mailAdapter = new ExchangeMailAdapter(exchangeService);
            var publisher = new SplunkAlertPublisher(splunkCollectorUrl, splunkApiKey);
            var alertGenerators = new IAlertGenerator[] { new VerifyAttachmentAlertGenerator(), new DefaultAlertGenerator() };
            var processor = new SplunkMailProcessor(mailAdapter, publisher, alertGenerators);

            var totalMessagesProcessed = 0;
            var continueProcessing = true;
            while (continueProcessing)
            {
                continueProcessing = false;
                Console.WriteLine("Checking Inbox");
                foreach (var item in mailAdapter.GetMailFrom(WellKnownFolderName.Inbox, exchangeMsgLimit))
                {
                    processor.ProcessMessage(item);
                    continueProcessing = true;
                    totalMessagesProcessed++;
                }
                Console.WriteLine("Checking Junk Folder");
                foreach (var item in mailAdapter.GetMailFrom(WellKnownFolderName.JunkEmail, exchangeMsgLimit))
                {
                    processor.ProcessMessage(item);
                    continueProcessing = true;
                    totalMessagesProcessed++;
                }
                Console.WriteLine("** Progress = {0} messages processed.", totalMessagesProcessed);
            }
            Console.WriteLine("\nProcessing Complete.");
        }
    }
}

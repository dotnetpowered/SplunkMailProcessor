using System;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;

namespace SplunkMailProcessor
{
    class Program
    {
        static void Main(string[] args)
        {   
            var splunkApiKey = ConfigurationManager.AppSettings["Splunk.ApiKey"];
            var splunkCollectorUrl = ConfigurationManager.AppSettings["Splunk.Url"];
            var exchangeUser = ConfigurationManager.AppSettings["Exchange.User"];
            var exchangePwd = ConfigurationManager.AppSettings["Exchange.Pwd"];
            var exchangeUrl = ConfigurationManager.AppSettings["Exchange.Url"];
            var exchangeMsgLimit = int.Parse(ConfigurationManager.AppSettings["Exchange.MessagesPerPoll"]);

            Console.WriteLine("Starting Splunk Mail Processor");
            Console.WriteLine("  Exchange Account: {0}", exchangeUser);
            Console.WriteLine("  Exchange Url: {0}", exchangeUrl);
            Console.WriteLine("  Splunk HTTP Event Collector Url: {0}", splunkCollectorUrl);

            var exchangeService = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Credentials = new WebCredentials(exchangeUser, exchangePwd),
                //service.TraceEnabled = true;
                Url = new Uri(exchangeUrl)
            };
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

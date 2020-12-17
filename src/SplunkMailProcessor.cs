using System;

namespace SplunkMailProcessor
{
    public class SplunkMailProcessor
    {
        ExchangeMailAdapter emailProcessor;
        SplunkAlertPublisher publisher;
        IAlertGenerator[] alertGenerators;

        public SplunkMailProcessor(ExchangeMailAdapter emailProcessor,
                 SplunkAlertPublisher publisher,
                 IAlertGenerator[] alertGenerators)
        {
            this.emailProcessor = emailProcessor;
            this.publisher = publisher;
            this.alertGenerators = alertGenerators;
        }

        public Alert ProcessMessage(MailItem item)
        {
            Alert alert = null;
            // Loop through alert generators - they will process sequentially 
            // First generator to create alert wins
            foreach (var alertGenerator in alertGenerators)
            {
                alert = alertGenerator.GetAlertFromMessage(item);
                if (alert != null)
                    break;  // alert was generated - so exit loop
            }

            if (alert != null)
            {
                Console.WriteLine("{0}: {1}", alert.timestamp, alert.message);
                publisher.CreateAlertAsync(alert).GetAwaiter().GetResult();
                Console.WriteLine(" Published alert to Splunk");
            }

            // Move mail message to archive folder
            emailProcessor.ArchiveMessage(item.ExchangeMailMessage);
            Console.WriteLine(" Message archived");

            return alert;
        }
    }
}

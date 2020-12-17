namespace SplunkMailProcessor
{
    public class DefaultAlertGenerator : IAlertGenerator
    {
        public Alert GetAlertFromMessage(MailItem mailItem)
        {
            return AlertFromMessage(mailItem);
        }

        public static Alert AlertFromMessage(MailItem mailItem)
        {
            var alert = new Alert()
            {
                id = mailItem.MessageId,
                message = mailItem.Subject,
                timestamp = string.Format("{0}",mailItem.Received),
                description = mailItem.Body,
                generatedBy = mailItem.From
            };

            return alert;
        }
    }
}

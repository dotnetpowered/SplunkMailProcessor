namespace SplunkMailProcessor
{
    /// <summary>
    /// Sample alert generator for handling attachments
    /// </summary>
    public class VerifyAttachmentAlertGenerator : IAlertGenerator
    {
        public Alert GetAlertFromMessage(MailItem mailItem)
        {
            if (string.IsNullOrWhiteSpace(mailItem.Subject))
            {
                return null;
            }

            var newSubject = mailItem.Subject;
            var createAlert = false;

            // Check the attachment content - only needed for specific messages
            if (newSubject.Contains("A specific message subject"))
            {
                newSubject += IsSuccessfulAttachment(mailItem, "Header text", "Success");
                
                createAlert = true;
            }
            if (createAlert)
            {
                var alert = DefaultAlertGenerator.AlertFromMessage(mailItem);
                alert.message = newSubject;
                return alert;
            }
            else
                return null;
        }

        string IsSuccessfulAttachment(MailItem mailItem, string fileHeader, string textToFind)
        {
            var result = "";
            foreach (var a in mailItem.Attachments)
            {
                if (a.Content.Contains(fileHeader))
                {
                    if (a.Content.Contains(textToFind))
                    {
                        result = "-" + fileHeader + "-success";
                    }
                    else
                    {
                        result = "-" + fileHeader + "-failed";
                    }
                }
            }
            return result;
        }
    }

 
}

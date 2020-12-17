using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace SplunkMailProcessor
{
    // Simplified mail item (with reference to original message)
    public class MailItem
    {
        public string MessageId { get; set; }
        public string From { get; set; }
        public string[] Recipients { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public IEnumerable<MailAttachment> Attachments { get; set; }
        public DateTime Received { get; set; }
        public EmailMessage ExchangeMailMessage;

        public static MailItem FromExchangeMailMessage(EmailMessage emailMessage)
        {
            var attachments = new List<MailAttachment>();
            if (emailMessage.HasAttachments)
            {
                foreach (var attachment in emailMessage.Attachments)
                {
                    if (attachment is FileAttachment fileAttachment)
                    {
                        fileAttachment.Load();
                        attachments.Add(new MailAttachment() {
                            Name = attachment.Name,
                            Content = UTF8Encoding.UTF8.GetString(fileAttachment.Content)
                        });
                    }
                }
            }
            var mailItem = new MailItem()
            {
                MessageId = emailMessage.Id.UniqueId,
                Received = emailMessage.DateTimeReceived,
                From = emailMessage.From.Address,
                Recipients = emailMessage.ToRecipients.Select(recipient => recipient.Address).ToArray(),
                Subject = emailMessage.Subject,
                Body = emailMessage.Body.ToString(),
                Attachments = attachments.ToArray(),
                ExchangeMailMessage = emailMessage
            };
            return mailItem;
        }
    }


}

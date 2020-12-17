using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace SplunkMailProcessor
{
    public class ExchangeMailAdapter
    {
        protected ExchangeService service;

        public ExchangeMailAdapter(ExchangeService service)
        {
            this.service = service;
        }

        public IEnumerable<MailItem> GetMailFrom(WellKnownFolderName folderName, int maxItems)
        {
            var view = new ItemView(maxItems)
            {
                Traversal = ItemTraversal.Shallow,
            };

            FindItemsResults<Item> findResults = service.FindItems(folderName, view);
            if (findResults.Count() > 0)
            {
                var itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, ItemSchema.Attachments)
                {
                    RequestedBodyType = BodyType.Text
                };
                ServiceResponseCollection<GetItemResponse> items =
                    service.BindToItems(findResults.Select(item => item.Id), itempropertyset);
                foreach (var item in items)
                {
                    var emailMessage = item.Item as EmailMessage;
                    var mailItem = MailItem.FromExchangeMailMessage(emailMessage);

                    yield return mailItem;
                }
            }
        }

        public void ArchiveMessage(EmailMessage message)
        {
            FindFoldersResults ArchiveFolder = service.FindFolders(WellKnownFolderName.MsgFolderRoot, new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Archive"), new FolderView(1));
            if (ArchiveFolder.Folders.Count == 1)
            {
                message.Move(ArchiveFolder.Folders[0].Id);
            }
        }
    }
}

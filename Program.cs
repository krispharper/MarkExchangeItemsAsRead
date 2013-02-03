using System;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace MarkExchangeItemsAsRead
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.AutodiscoverUrl(UserPrincipal.Current.EmailAddress);
            Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);

            FolderView folderView = new FolderView(inbox.ChildFolderCount);
            folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.ChildFolderCount);
            SearchFilter filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Notifications");

            Folder folder = inbox.FindFolders(filter, folderView).First();
            MarkAllAsRead(folder);
        }

        static void MarkAllAsRead(Folder folder)
        {
            if (folder.ChildFolderCount == 0)
            {
                ItemView itemView = new ItemView(100);
                itemView.PropertySet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.IsRead);
                SearchFilter filter = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
                folder.FindItems(filter, itemView)
                      .Cast<EmailMessage>()
                      .ToList()
                      .ForEach(item =>
                {
                    item.IsRead = true;
                    item.Update(ConflictResolutionMode.AutoResolve);
                });
            }
            else
            {
                FolderView folderView = new FolderView(folder.ChildFolderCount);
                folder.FindFolders(folderView).ToList().ForEach(child => MarkAllAsRead(child));
            }
        }
    }
}

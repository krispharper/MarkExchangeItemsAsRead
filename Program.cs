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
            Folder root = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);
            FolderView folderView = new FolderView(10);
            SearchFilter filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Archived");
            FindFoldersResults results = service.FindFolders(WellKnownFolderName.MsgFolderRoot, filter, folderView);
            MarkAllAsRead(results.First());
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

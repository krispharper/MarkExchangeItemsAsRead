using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MarkAsRead
{
    class Program
    {
        static void Main(string[] args)
        {
            var stream = new FileStream("./log.txt", FileMode.OpenOrCreate, FileAccess.Write);
            var writer = new StreamWriter(stream);
            Console.SetOut(writer);

            try
            {
                Outlook.Application outlook = new Outlook.Application();
                Outlook.MAPIFolder root = outlook.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent;
                Outlook.Folders folders = root.Folders;

                Console.WriteLine("Starting to mark items as read");

                for (int i = 1; i < folders.Count; i++)
                {
                    Outlook.MAPIFolder folder = root.Folders[i];

                    if (folder.Name == "Archived" || folder.Name == "Deleted Items")
                    {
                        Outlook.Items items = folder.Items;

                        Outlook.MailItem item = items.Find("[Unread] = true");

                        while (item != null)
                        {
                            Console.WriteLine($"Marking item with subject {item.Subject} as read");
                            item.UnRead = false;
                            item.Save();
                            Marshal.ReleaseComObject(item);
                            item = items.FindNext();
                        }

                        Marshal.ReleaseComObject(items);
                        items = null;
                    }

                    Marshal.ReleaseComObject(folder);
                    folder = null;
                }

                Marshal.ReleaseComObject(folders);
                folders = null;

                Marshal.ReleaseComObject(root);
                root = null;

                Marshal.ReleaseComObject(outlook);
                outlook = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                Console.WriteLine("Finished marking items as read");
                writer.Close();
                stream.Close();
            }
        }
    }
}

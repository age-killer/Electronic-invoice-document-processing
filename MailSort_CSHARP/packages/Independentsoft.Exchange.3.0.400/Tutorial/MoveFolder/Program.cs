using System;
using System.Net;
using Independentsoft.Exchange;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            Service service = new Service("https://myserver/ews/Exchange.asmx", credential);

            try
            {
                IsEqualTo restriction = new IsEqualTo(FolderPropertyPath.DisplayName, "Test");

                FindFolderResponse findFolderResponse = service.FindFolder(StandardFolder.MailboxRoot, restriction);

                for (int i = 0; i < findFolderResponse.Folders.Count; i++)
                {
                    FolderId currentFolderId = findFolderResponse.Folders[i].FolderId;

                    FolderId newFolderId = service.MoveFolder(currentFolderId, StandardFolder.Drafts);
                }
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("Error: " + ex.XmlMessage);
                Console.Read();
            }
            catch (WebException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.Read();
            }
        }
    }
}

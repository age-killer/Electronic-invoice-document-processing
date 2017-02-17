using System;
using System.Net;
using System.Collections.Generic;
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
                //Initial synscronization. It will recursively retrieve all existing folders from the mailbox 
                SyncFoldersResponse response = service.SyncFolders(StandardFolder.MailboxRoot);

                //Keep synchronization state in order to use it in next synchronizations
                string state = response.State;

                IList<Folder> allExistingFolders = response.CreatedFolders;

                //Create new folder
                FolderId folderId = service.CreateFolder("test", StandardFolder.MailboxRoot);

                //Synchronize folders again but now include the state parameter in order to find changes from last synchronization
                SyncFoldersResponse newResponse = service.SyncFolders(StandardFolder.MailboxRoot, state);

                IList<Folder> newFolders = newResponse.CreatedFolders;
                IList<FolderId> deletedFolders = newResponse.DeletedFolders;
                IList<Folder> updatedFolders = newResponse.UpdatedFolders;

                //Display new created folders
                for (int i = 0; i < newFolders.Count; i++)
                {
                    Console.WriteLine(newFolders[i].DisplayName);
                    Console.WriteLine(newFolders[i].FolderId);
                }

                Console.Read();
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

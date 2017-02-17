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
                IsGreaterThanOrEqualTo restriction1 = new IsGreaterThanOrEqualTo(FolderPropertyPath.CreationTime, DateTime.Today);
                IsLessThan restriction2 = new IsLessThan(FolderPropertyPath.CreationTime, DateTime.Today.AddDays(1));
                And restriction3 = new And(restriction1, restriction2);
               
                FindFolderResponse findFolderResponse = service.FindFolder(StandardFolder.MailboxRoot, FolderPropertyPath.AllPropertyPaths, restriction3, FolderQueryTraversal.Deep);

                for (int i = 0; i < findFolderResponse.Folders.Count; i++)
                {
                    Console.WriteLine(findFolderResponse.Folders[i].DisplayName);
                    Console.WriteLine(findFolderResponse.Folders[i].FolderClass);
                    Console.WriteLine(findFolderResponse.Folders[i].CreationTime);
                    Console.WriteLine(findFolderResponse.Folders[i].FolderId);
                    Console.WriteLine("----------------------------------------------------------");
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

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
                Folder folder1 = new Folder("Test1");
                Folder folder2 = new Folder("Test2");

                FolderId folder1Id = service.CreateFolder(folder1, StandardFolder.Drafts);
                FolderId folder2Id = service.CreateFolder(folder2, StandardFolder.Drafts);

                Response response1 = service.DeleteFolder(folder1Id); //hard delete
                Response response2 = service.DeleteFolder(folder2Id, DeleteType.MoveToDeletedItems);
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

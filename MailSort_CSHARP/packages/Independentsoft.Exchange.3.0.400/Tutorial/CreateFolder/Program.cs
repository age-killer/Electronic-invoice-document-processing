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

                //Create a folder in the mailbox root
                FolderId folder1Id = service.CreateFolder(folder1, StandardFolder.MailboxRoot);

                //Create a subfolder of the Test1
                FolderId folder2Id = service.CreateFolder(folder2, folder1Id);

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

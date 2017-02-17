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

            //works on Exchange 2013 / Office365 and higher
            service.RequestServerVersion = RequestServerVersion.Exchange2013;

            try
            {
                IList<Folder> subfolders = new List<Folder>();
                subfolders.Add(new Folder("Folder1")); //Folder1 is under Inbox folder
                subfolders.Add(new Folder("Folder2")); //Folder2 is under Folder1
                subfolders.Add(new Folder("Folder3")); //Folder3 is under Folder2

                IList<FolderInfoResponse> response = service.CreateFolderPath(subfolders, StandardFolder.Inbox);
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

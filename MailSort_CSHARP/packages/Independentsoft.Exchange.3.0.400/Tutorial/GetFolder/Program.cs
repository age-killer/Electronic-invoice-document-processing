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
                //Get all folder properties
                Folder inboxFolder1 = service.GetFolder(StandardFolder.Inbox);

                //Get only FolderId property of the Inbox folder
                Folder inboxFolder2 = service.GetFolder(StandardFolder.Inbox, ShapeType.Id);

                IList<PropertyPath> propertyPaths = new List<PropertyPath>();
                propertyPaths.Add(FolderPropertyPath.DisplayName);
                propertyPaths.Add(FolderPropertyPath.FolderClass);
                propertyPaths.Add(FolderPropertyPath.Comment);

                //Get specified properties
                Folder inboxFolder3 = service.GetFolder(StandardFolder.Inbox, propertyPaths);
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

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
                PropertyName myCustomPropertyName = new PropertyName("myproperty", StandardPropertySet.PublicStrings, MapiPropertyType.String);

                Property myCustomProperty = new Property(myCustomPropertyName, "value1");
                Property commentProperty = new Property(FolderPropertyPath.Comment, "Folder description text.");

                //Get only FolderId property of the Inbox folder
                Folder inboxFolder = service.GetFolder(StandardFolder.Inbox, ShapeType.Id);

                FolderId newFolderId1 = service.UpdateFolder(inboxFolder.FolderId, myCustomProperty);
                FolderId newFolderId2 = service.UpdateFolder(inboxFolder.FolderId, commentProperty);

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

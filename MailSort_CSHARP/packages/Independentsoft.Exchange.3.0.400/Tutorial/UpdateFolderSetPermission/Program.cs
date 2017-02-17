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
                Permission permission1 = new Permission();
                permission1.UserId = new UserId("Mark@mydomain.com");

                permission1.Level = PermissionLevel.Custom;
                permission1.CanCreateItems = true;
                permission1.CanCreateSubFolders = true;

                PermissionSet permissionSet = new PermissionSet();
                permissionSet.Permissions.Add(permission1);

                Property permissionSetProperty = new Property(FolderPropertyPath.PermissionSet, permissionSet);

                Folder contactsFolder = service.GetFolder(StandardFolder.Contacts);

                FolderChange change = new FolderChange(contactsFolder.FolderId, permissionSetProperty);

                service.UpdateFolder(change);
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

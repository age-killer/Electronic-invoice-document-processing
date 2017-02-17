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
                CalendarPermission permission1 = new CalendarPermission();
                permission1.UserId = new UserId("Mark@mydomain.com");
                permission1.Level = CalendarPermissionLevel.Author;

                CalendarFolder calFolder = new CalendarFolder("TestCalendar");
                calFolder.PermissionSet = new CalendarPermissionSet();
                calFolder.PermissionSet.Permissions.Add(permission1);

                FolderId folder1Id = service.CreateFolder(calFolder, StandardFolder.Calendar);
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

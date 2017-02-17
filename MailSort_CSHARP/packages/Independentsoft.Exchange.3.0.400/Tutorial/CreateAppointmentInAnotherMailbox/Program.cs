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
                Mailbox johnMailbox = new Mailbox("John@mydomain.com");
                StandardFolderId johnCalendarFolder = new StandardFolderId(StandardFolder.Calendar, johnMailbox);
                
                Appointment appointment = new Appointment();
                appointment.Subject = "Test";
                appointment.Body = new Body("Body text");
                appointment.StartTime = DateTime.Today.AddHours(15);
                appointment.EndTime = DateTime.Today.AddHours(16);
                appointment.Location = "My Office";
                appointment.ReminderIsSet = true;
                appointment.ReminderMinutesBeforeStart = 30;

                ItemId itemId = service.CreateItem(appointment, johnCalendarFolder);
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

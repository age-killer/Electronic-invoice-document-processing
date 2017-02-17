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

                FindItemResponse response = service.FindItem(johnCalendarFolder, AppointmentPropertyPath.AllPropertyPaths);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Appointment)
                    {
                        Appointment appointment = (Appointment)response.Items[i];

                        Console.WriteLine("Subject = " + appointment.Subject);
                        Console.WriteLine("StartTime = " + appointment.StartTime);
                        Console.WriteLine("EndTime = " + appointment.EndTime);
                        Console.WriteLine("Body Preview = " + appointment.BodyPlainText);
                        Console.WriteLine("----------------------------------------------------------------");
                    }
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

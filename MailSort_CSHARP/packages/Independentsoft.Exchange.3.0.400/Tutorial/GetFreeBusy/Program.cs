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
                MailboxData mailbox = new MailboxData("John@mydomain.com");
                SerializableTimeZone timeZone = new SerializableTimeZone(StandardTimeZone.Berlin);
                FreeBusyViewOptions freeBusyOptions = new FreeBusyViewOptions(DateTime.Today, DateTime.Today.AddDays(10), FreeBusyViewType.DetailedMerged);

                AvailabilityResponse response = service.GetAvailability(mailbox, timeZone, freeBusyOptions);

                for (int i = 0; i < response.FreeBusyResponses.Count; i++)
                {
                    FreeBusyView freeBusyView = response.FreeBusyResponses[i].FreeBusyView;

                    for (int j = 0; j < freeBusyView.CalendarEvents.Count; j++)
                    {
                        Console.WriteLine("Subject = " + freeBusyView.CalendarEvents[j].Subject);
                        Console.WriteLine("Location = " + freeBusyView.CalendarEvents[j].Location);
                        Console.WriteLine("StartTime = " + freeBusyView.CalendarEvents[j].StartTime);
                        Console.WriteLine("EndTime = " + freeBusyView.CalendarEvents[j].EndTime);
                        Console.WriteLine("-----------------------------------------------------");
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

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

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1;

            try
            {
                CalendarOptions options = service.GetCalendarOptions();

                if (options != null)
                {
                    Console.WriteLine("AutoProcess:" + options.AutoProcess);
                    Console.WriteLine("AutoProcessExternal:" + options.AutoProcessExternal);
                    Console.WriteLine("AutoProcessNewItems:" + options.AutoProcessNewItems);
                    Console.WriteLine("DeleteUpdatedItems:" + options.DeleteUpdatedItems);
                    Console.WriteLine("ReminderDefault:" + options.ReminderDefault);
                    Console.WriteLine("ReminderUpgradeTime:" + options.ReminderUpgradeTime);
                    Console.WriteLine("RemoveForwardedMeetingNotifications:" + options.RemoveForwardedMeetingNotifications);
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

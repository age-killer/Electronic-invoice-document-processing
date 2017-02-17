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

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1;

            try
            {
                TimeZoneDefinition timeZone = GetTimeZone(service, "Berlin"); //find time zone based on name
                //TimeZoneDefinition timeZone = GetTimeZone(service, "UTC-08:00"); //find time zone based on offset

                Appointment appointment = new Appointment();
                appointment.Subject = "Test";
                appointment.Body = new Body("Body text");
                appointment.StartTime = DateTime.Today;
                appointment.EndTime = DateTime.Today.AddDays(1);
                appointment.Location = "My Office";
                appointment.IsAllDayEvent = true;
                appointment.StartTimeZone = timeZone;
                appointment.EndTimeZone = timeZone;

                ItemId itemId = service.CreateItem(appointment);
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

        private static TimeZoneDefinition GetTimeZone(Service service, string name)
        {
            GetServerTimeZonesResponse response = service.GetServerTimeZones();

            foreach (TimeZoneDefinition timeZone in response.TimeZoneDefinitions)
            {
                if (timeZone.Name.IndexOf(name) > 0)
                {
                    return timeZone;
                }
            }

            return null;
        }
    }
}

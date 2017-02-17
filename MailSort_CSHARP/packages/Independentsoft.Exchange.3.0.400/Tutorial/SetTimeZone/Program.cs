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
                //Set MeetingTimeZone property. Use only on Exchange 2007
                Appointment appointment1 = new Appointment();
                appointment1.Subject = "Test1";
                appointment1.Body = new Body("Body text");
                appointment1.StartTime = DateTime.Today.AddHours(10);
                appointment1.EndTime = DateTime.Today.AddHours(11);
                appointment1.MeetingTimeZone = new Independentsoft.Exchange.TimeZone(StandardTimeZone.Berlin);

                ItemId itemId1 = service.CreateItem(appointment1);

                service.RequestServerVersion = RequestServerVersion.Exchange2010SP1;

                TimeZoneDefinition timeZone = GetTimeZone(service, "Berlin"); //find time zone based on name
                //TimeZoneDefinition timeZone = GetTimeZone(service, "UTC-08:00"); //find time zone based on offset

                //Setting StartTimeZone and EndTimeZone properties. Works on Exchange 2010 and later
                Appointment appointment2 = new Appointment();
                appointment2.Subject = "Test2";
                appointment2.Body = new Body("Body text");
                appointment2.StartTime = DateTime.Today.AddHours(12);
                appointment2.EndTime = DateTime.Today.AddHours(13);
                appointment2.StartTimeZone = timeZone;
                appointment2.EndTimeZone = timeZone;

                ItemId itemId2 = service.CreateItem(appointment2);

                //Setting TimeZoneContext. It is time zone for entire session
                service.TimeZoneContext = timeZone;

                Appointment appointment3 = new Appointment();
                appointment3.Subject = "Test3";
                appointment3.Body = new Body("Body text");
                appointment3.StartTime = DateTime.Today.AddHours(15);
                appointment3.EndTime = DateTime.Today.AddHours(16);

                ItemId itemId3 = service.CreateItem(appointment3);
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

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
                GetServerTimeZonesResponse response = service.GetServerTimeZones();

                //Display all available time zones
                foreach (TimeZoneDefinition timeZone in response.TimeZoneDefinitions)
                {
                    Console.WriteLine(timeZone.Id);
                    Console.WriteLine(timeZone.Name);
                    Console.WriteLine("-------------------------------------------------------");
                }

                //Display time zones with offset -05:00
                foreach (TimeZoneDefinition timeZone in response.TimeZoneDefinitions)
                {
                    if (timeZone.Name.IndexOf("UTC-05:00") > 0)
                    {
                        Console.WriteLine(timeZone.Id);
                        Console.WriteLine(timeZone.Name);
                        Console.WriteLine("-------------------------------------------------------");
                    }
                }

                //Find time zone based on name
                foreach (TimeZoneDefinition timeZone in response.TimeZoneDefinitions)
                {
                    if (timeZone.Name.IndexOf("Dublin") > 0)
                    {
                        Console.WriteLine(timeZone.Id);
                        Console.WriteLine(timeZone.Name);
                        Console.WriteLine("-------------------------------------------------------");
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

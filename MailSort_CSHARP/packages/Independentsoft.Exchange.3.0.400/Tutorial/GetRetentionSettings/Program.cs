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
                RetentionSettings settings = service.GetRetentionSettings(StandardFolder.Inbox);

                foreach (RetentionPolicyTag tag in settings.RetentionPolicyTags)
                {
                    Console.WriteLine(tag.Id);
                    Console.WriteLine(tag.DisplayName);
                    Console.WriteLine(tag.RetentionAction);
                    Console.WriteLine(tag.Period);
                    Console.WriteLine(tag.IsArchive);
                    Console.WriteLine(tag.IsVisible);
                    Console.WriteLine("---------------------------------------------");
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

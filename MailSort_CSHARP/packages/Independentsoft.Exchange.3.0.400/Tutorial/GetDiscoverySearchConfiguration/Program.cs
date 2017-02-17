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

            service.RequestServerVersion = RequestServerVersion.Exchange2013;

            try
            {
                GetDiscoverySearchConfigurationResponse response = service.GetDiscoverySearchConfiguration("MyDiscoverySearchTest");

                foreach (DiscoverySearchConfiguration configuration in response.DiscoverySearchConfigurations)
                {
                    foreach (SearchableMailbox mailbox in configuration.SearchableMailboxes)
                    {
                        Console.WriteLine("DisplayName:" + mailbox.DisplayName);
                        Console.WriteLine("PrimarySmtpAddress:" + mailbox.PrimarySmtpAddress);
                        Console.WriteLine("ReferenceId:" + mailbox.ReferenceId);
                        Console.WriteLine("--------------------------------------------");
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

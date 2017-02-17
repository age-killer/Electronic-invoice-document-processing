using System;
using System.Net;
using Independentsoft.Exchange;
using Independentsoft.Exchange.Autodiscover;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            AutodiscoverService autodiscoverService = new AutodiscoverService("https://myserver/autodiscover/autodiscover.xml", credential);

            try
            {
                MobileSyncProvider mobileSyncProvider = autodiscoverService.AutodiscoverMobileSyncProvider("John@mydomain.com");

                Console.WriteLine("Culture = " + mobileSyncProvider.Culture);

                if (mobileSyncProvider.User != null)
                {
                    MobileSyncUser user = mobileSyncProvider.User;

                    Console.WriteLine("DisplayName = " + user.DisplayName);
                    Console.WriteLine("EmailAddress = " + user.EmailAddress);
                }

                if (mobileSyncProvider.Server != null)
                {
                    MobileSyncServer server = mobileSyncProvider.Server;

                    Console.WriteLine("Server Name = " + server.Name);
                    Console.WriteLine("Server Type = " + server.Type);
                    Console.WriteLine("Server Url  = " + server.Url);
                }

                Console.Read();
            }
            catch (AutodiscoverException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
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

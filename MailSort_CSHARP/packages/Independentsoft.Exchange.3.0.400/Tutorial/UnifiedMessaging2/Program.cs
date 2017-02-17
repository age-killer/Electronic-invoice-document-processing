using System;
using System.Net;
using Independentsoft.Exchange;
using Independentsoft.Exchange.UnifiedMessaging;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            UnifiedMessagingService unifiedMessagingService = new UnifiedMessagingService("https://myserver/UnifiedMessaging/service.asmx", credential);

            try
            {
                string response = unifiedMessagingService.PlayOnPhone("AAAAAGsd2rbQLVtLobUGbrq/9IUHAEX2ikn/L8JJtI5WHI0FAW8AAAFXHhsAACxVpEl+KVVLl957wp//x6UAGAetcDUAAA==", "12345");

                Console.WriteLine(response);
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

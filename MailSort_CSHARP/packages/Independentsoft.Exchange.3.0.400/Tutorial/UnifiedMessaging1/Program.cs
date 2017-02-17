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
                bool enabled = unifiedMessagingService.IsUnifiedMessagingEnabled();
                UnifiedMessagingProperties properties = unifiedMessagingService.GetUnifiedMessagingProperties();

                Console.WriteLine("Enabled = " + enabled);

                if (properties != null)
                {
                    Console.WriteLine("OutOfOfficeEnabled = " + properties.OutOfOfficeEnabled);
                    Console.WriteLine("MissedCallNotificationEnabled = " + properties.MissedCallNotificationEnabled);
                    Console.WriteLine("PlayOnPhoneDialString = " + properties.PlayOnPhoneDialString);
                    Console.WriteLine("TelephoneAccessFolderEmail = " + properties.TelephoneAccessFolderEmail);
                    Console.WriteLine("TelephoneAccessNumbers = " + properties.TelephoneAccessNumbers);
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

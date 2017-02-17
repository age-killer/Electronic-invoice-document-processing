using System;
using System.Net;
using Independentsoft.Exchange;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "passsword");
            Service service = new Service("https://myserver/ews/Exchange.asmx", credential);

            try
            {
                Mailbox mailbox = new Mailbox("John@mydomain.com");
                mailbox.RoutingType = "SMTP";

                ServiceConfigurationResponse response1 = service.GetServiceConfiguration(mailbox, ServiceConfigurationType.MailTips);

                Console.WriteLine("IsMailTipsEnabled=" + response1.MailTipsConfiguration.IsMailTipsEnabled);
                Console.WriteLine("LargeAudienceThreshold=" + response1.MailTipsConfiguration.LargeAudienceThreshold);
                Console.WriteLine("MaxMessageSize=" + response1.MailTipsConfiguration.MaxMessageSize);
                Console.WriteLine("MaxRecipients=" + response1.MailTipsConfiguration.MaxRecipients);
                Console.WriteLine("ShowExternalRecipientCount=" + response1.MailTipsConfiguration.ShowExternalRecipientCount);

                ServiceConfigurationResponse response2 = service.GetServiceConfiguration(ServiceConfigurationType.UnifiedMessagingConfiguration);

                Console.WriteLine("IsUnifiedMessagingEnabled=" + response2.UnifiedMessagingConfiguration.IsUnifiedMessagingEnabled);
                Console.WriteLine("IsPlayOnPhoneEnabled=" + response2.UnifiedMessagingConfiguration.IsPlayOnPhoneEnabled);
                Console.WriteLine("PlayOnPhoneDialString=" + response2.UnifiedMessagingConfiguration.PlayOnPhoneDialString);

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

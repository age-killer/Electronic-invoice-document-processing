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
                Mailbox mailbox = new Mailbox("John@mydomain.com");

                FindMessageTrackingReportResponse response = service.FindMessageTrackingReport(MessageTrackingScope.Organization, "mydomain.com", mailbox);

                foreach (MessageTrackingSearchResult result in response.MessageTrackingSearchResults)
                {
                    Console.WriteLine("Subject=" + result.Subject);

                    if (result.Sender != null)
                    {
                        Console.WriteLine("Sender=" + result.Sender.EmailAddress);
                    }

                    if (result.Sender != null)
                    {
                        Console.WriteLine("PurportedSender=" + result.PurportedSender.EmailAddress);
                    }

                    Console.WriteLine("PreviousHopServer=" + result.PreviousHopServer);
                    Console.WriteLine("MessageTrackingReportId=" + result.MessageTrackingReportId);
                    Console.WriteLine("SubmittedTime=" + result.SubmittedTime);

                    foreach (Mailbox recipient in result.Recipients)
                    {
                        Console.WriteLine("Recipient=" + recipient.EmailAddress);
                    }

                    Console.WriteLine("-----------------------------------------------");
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

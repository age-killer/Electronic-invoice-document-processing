using System;
using System.IO;
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
                StreamingSubscription subscription = new StreamingSubscription(StandardFolder.Inbox, EventType.NewMail);

                SubscribeResponse response = service.Subscribe(subscription);

                Stream responseStream = service.GetStreamingEvents(response.SubscriptionId, 30);

                while (true)
                {
                    StreamingEventsResponse eventsResponse = new StreamingEventsResponse(responseStream);

                    Console.WriteLine("New mails:" + eventsResponse.Notifications.Count);

                    //exit when connection is closed or error occurs.
                    if (eventsResponse.ConnectionStatus == ConnectionStatus.Closed)
                    {
                        break;
                    }
                    else if (eventsResponse.ResponseClass == ResponseClass.Error)
                    {
                        Console.WriteLine(eventsResponse.XmlMessage);
                        Console.WriteLine(eventsResponse.Message);
                        break;
                    }
                }
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

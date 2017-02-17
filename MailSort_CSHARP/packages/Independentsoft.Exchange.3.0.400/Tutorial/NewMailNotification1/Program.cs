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
                string newWatermark = null;

                PullSubscription subscription = new PullSubscription(StandardFolder.Inbox, EventType.NewMail);

                //initail subscribe
                SubscribeResponse subscribeResponse = service.Subscribe(subscription);

                while (true)
                {
                    //wait 60 seconds
                    System.Threading.Thread.CurrentThread.Join(60000);

                    GetEventsResponse eventsResponse = service.GetEvents(subscribeResponse);

                    Notification notification = eventsResponse.Notification;

                    for (int i = 0; i < notification.Events.Count; i++)
                    {
                        newWatermark = notification.Events[i].Watermark;

                        if (notification.Events[i] is NewMailEvent)
                        {
                            NewMailEvent newMailEvent = (NewMailEvent)notification.Events[i];

                            ItemId itemId = (ItemId)newMailEvent.Id;

                            Message message = service.GetMessage(itemId);

                            Console.WriteLine(message.Subject);
                            Console.WriteLine(message.ReceivedTime);
                        }
                    }

                    //resubscribe with new watermark 
                    subscription.Watermark = newWatermark;
                    subscribeResponse = service.Subscribe(subscription);
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

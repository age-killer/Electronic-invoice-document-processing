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

            try
            {
                string newWatermark = null;

                IList<EventType> eventTypes = new List<EventType>();
                eventTypes.Add(EventType.Created);
                eventTypes.Add(EventType.Modified);

                PullSubscription subscription = new PullSubscription(StandardFolder.Calendar, eventTypes);

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

                        if (notification.Events[i] is CreatedEvent)
                        {
                            CreatedEvent createdEvent = (CreatedEvent)notification.Events[i];

                            if (createdEvent.Id is ItemId)
                            {
                                ItemId itemId = (ItemId)createdEvent.Id;
                                Console.WriteLine("Created item = " + itemId);
                            }
                            else if (createdEvent.Id is FolderId)
                            {
                                FolderId folderId = (FolderId)createdEvent.Id;
                                Console.WriteLine("Created folder = " + folderId);
                            }
                        }
                        else if (notification.Events[i] is ModifiedEvent)
                        {
                            ModifiedEvent modifiedEvent = (ModifiedEvent)notification.Events[i];

                            if (modifiedEvent.Id is ItemId)
                            {
                                ItemId itemId = (ItemId)modifiedEvent.Id;
                                Console.WriteLine("Modified item = " + itemId);
                            }
                            else if (modifiedEvent.Id is FolderId)
                            {
                                FolderId folderId = (FolderId)modifiedEvent.Id;
                                Console.WriteLine("Modified folder = " + folderId);
                            }
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

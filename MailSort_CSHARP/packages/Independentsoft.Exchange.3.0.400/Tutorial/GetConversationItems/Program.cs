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

            service.RequestServerVersion = RequestServerVersion.Exchange2013;

            try
            {
                IsGreaterThanOrEqualTo restriction = new IsGreaterThanOrEqualTo(MessagePropertyPath.ReceivedTime, DateTime.Now.AddMinutes(-15));

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, new ItemShape(ShapeType.Id), restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        ConversationRequest request1 = new ConversationRequest(response.Items[i].ItemId);

                        IList<ConversationItemResponse> conversations = service.GetConversationItems(request1);

                        foreach (ConversationItemResponse conversation in conversations)
                        {
                            foreach (ConversationNode node in conversation.Nodes)
                            {
                                foreach (Item item in node.Items)
                                {
                                    if (item is Message)
                                    {
                                        Message message = (Message)item;

                                        Console.WriteLine("Subject:" + message.Subject);
                                        Console.WriteLine("ReceivedTime:" + message.ReceivedTime);
                                        Console.WriteLine("---------------------------------------------------------");
                                    }
                                }
                            }
                            
                        }
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

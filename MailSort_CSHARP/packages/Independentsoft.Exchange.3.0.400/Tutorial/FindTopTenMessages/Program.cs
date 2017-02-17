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
                IList<PropertyOrder> order = new List<PropertyOrder>();
                order.Add(new PropertyOrder(MessagePropertyPath.ReceivedTime, SortDirection.Descending));

                IndexedPageView view = new IndexedPageView();
                view.MaxEntriesReturned = 10;

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, order, view);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        Message message = (Message)response.Items[i];

                        Console.WriteLine("Subject = " + message.Subject);
                        Console.WriteLine("ReceivedTime = " + message.ReceivedTime);

                        if (message.From != null)
                        {
                            Console.WriteLine("From = " + message.From.Name);
                        }

                        Console.WriteLine("Body Preview = " + message.BodyPlainText);
                        Console.WriteLine("----------------------------------------------------------------");
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

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
                int offset = 0;

                while (true)
                {
                    IndexedPageView view = new IndexedPageView(offset, IndexBasePoint.Beginning, 100);

                    FindItemResponse response = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, view);

                    for (int i = 0; i < response.Items.Count; i++)
                    {
                        if (response.Items[i] is Message)
                        {
                            Message message = (Message)response.Items[i];

                            Console.WriteLine("Subject = " + message.Subject);
                            Console.WriteLine("ReceivedTime = " + message.ReceivedTime);
                            Console.WriteLine("----------------------------------------------------------------");
                        }
                    }

                    if (response.IndexedPagingOffset < response.TotalItemsInView)
                    {
                        offset = response.IndexedPagingOffset;
                    }
                    else
                    {
                        break;
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

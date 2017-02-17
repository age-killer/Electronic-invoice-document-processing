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
                IList<ItemId> items = new List<ItemId>();

                IsLessThanOrEqualTo restriction = new IsLessThanOrEqualTo(MessagePropertyPath.ReceivedTime, DateTime.Now.AddMonths(-1));

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, new ItemShape(ShapeType.Id), restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        items.Add(response.Items[i].ItemId);
                    }
                }

                if (items.Count > 0)
                {
                    IList<ItemInfoResponse> archiveResponse = service.ArchiveItem(items, StandardFolder.Inbox);
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

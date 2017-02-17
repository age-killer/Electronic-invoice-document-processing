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
                ItemShape itemShape = new ItemShape(ShapeType.Id);
                IsEqualTo restriction = new IsEqualTo(ItemPropertyPath.ItemClass, ItemClass.MeetingRequest);

                FindItemResponse findItemResponse = service.FindItem(StandardFolder.Inbox, itemShape, restriction);

                foreach(Item item in findItemResponse.Items)
                {
                    AcceptItem acceptItem = new AcceptItem(item.ItemId);

                    ItemInfoResponse response = service.AcceptMeetingRequest(acceptItem);
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

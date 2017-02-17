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
                IsEqualTo restriction = new IsEqualTo(AppointmentPropertyPath.Subject, "Meeting");

                FindItemResponse findItemResponse = service.FindItem(StandardFolder.Calendar, restriction);

                if (findItemResponse.Items.Count == 1)
                {
                    ItemId itemId = findItemResponse.Items[0].ItemId;

                    CancelItem cancelItem = new CancelItem(itemId);

                    ItemInfoResponse response = service.CancelMeetingRequest(cancelItem);
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

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
                Contains restriction = new Contains(MessagePropertyPath.Subject, "test", ContainmentMode.ExactPhrase, ContainmentComparison.IgnoreCase);

                FindItemResponse response = service.FindItem(StandardFolder.Calendar, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Appointment)
                    {
                        Property requiredAttendees = new Property(AppointmentPropertyPath.RequiredAttendees, new Attendee("John@mydomain.com"));

                        ItemChange itemChange = new ItemChange();
                        itemChange.ItemId = response.Items[i].ItemId;
                        itemChange.PropertiesToAppend.Add(requiredAttendees);

                        ItemId newItemId = service.UpdateItem(itemChange, SendMeetingOption.SendToAll);
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

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
                IsEqualTo restriction = new IsEqualTo(AppointmentPropertyPath.Subject, "Meeting");

                FindItemResponse findItemResponse = service.FindItem(StandardFolder.Calendar, restriction);

                Property startTimeProperty = new Property(AppointmentPropertyPath.StartTime, DateTime.Today.AddHours(17));
                Property endTimeProperty = new Property(AppointmentPropertyPath.EndTime, DateTime.Today.AddHours(20));

                IList<Property> properties = new List<Property>();
                properties.Add(startTimeProperty);
                properties.Add(endTimeProperty);

                if (findItemResponse.Items.Count == 1)
                {
                    ItemId itemId = findItemResponse.Items[0].ItemId;

                    itemId = service.UpdateItem(itemId, properties, SendMeetingOption.SendToAllAndSaveCopy);
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

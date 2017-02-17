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
                IsEqualTo restriction1 = new IsEqualTo(AppointmentPropertyPath.StartTime, DateTime.Today.AddHours(15));
                IsEqualTo restriction2 = new IsEqualTo(AppointmentPropertyPath.EndTime, DateTime.Today.AddHours(16));
                And restriction3 = new And(restriction1, restriction2);

                FindItemResponse response = service.FindItem(StandardFolder.Calendar, restriction3);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Appointment)
                    {
                        ItemId itemId = response.Items[i].ItemId;

                        Property startTimeProperty = new Property(AppointmentPropertyPath.StartTime, DateTime.Today.AddHours(18));
                        Property endTimeProperty = new Property(AppointmentPropertyPath.EndTime, DateTime.Today.AddHours(19));

                        IList<Property> properties = new List<Property>();
                        properties.Add(startTimeProperty);
                        properties.Add(endTimeProperty);

                        itemId = service.UpdateItem(itemId, properties);
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


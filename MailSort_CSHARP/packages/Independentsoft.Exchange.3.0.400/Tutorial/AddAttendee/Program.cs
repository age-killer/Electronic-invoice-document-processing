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
                Appointment appointment = service.GetAppointment("AAMkAAAAAAAAYAABClH/p3xPfSLGfvkhFmBB/AAvTWJunAAA=");

                ItemId itemId = appointment.ItemId;

                ItemChange itemChange = new ItemChange(itemId);
                itemChange.PropertiesToAppend.Add(new Property(AppointmentPropertyPath.RequiredAttendees, new Attendee("John@mydomain.com")));

                itemId = service.UpdateItem(itemChange, SendMeetingOption.SendToChanged);
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

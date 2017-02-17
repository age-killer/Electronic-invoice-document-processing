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

                //remove
                ItemChange itemChange0 = new ItemChange(itemId);
                itemChange0.PropertiesToDelete.Add(AppointmentPropertyPath.RequiredAttendees);
                itemChange0.PropertiesToDelete.Add(AppointmentPropertyPath.OptionalAttendees);

                itemId = service.UpdateItem(itemChange0);

                //add first
                ItemChange itemChange1 = new ItemChange(itemId);
                itemChange1.PropertiesToSet.Add(new Property(AppointmentPropertyPath.RequiredAttendees, new Attendee("John@mydomain.com")));

                itemId = service.UpdateItem(itemChange1);

                //append others 
                ItemChange itemChange2 = new ItemChange(itemId);
                itemChange2.PropertiesToAppend.Add(new Property(AppointmentPropertyPath.RequiredAttendees, new Attendee("Peter@mydomain.com")));

                //send meeting update.
                itemId = service.UpdateItem(itemChange2, SendMeetingOption.SendToAll);
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

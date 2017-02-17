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
                Appointment appointment = new Appointment();
                appointment.Subject = "Meeting";
                appointment.Body = new Body("Body text");
                appointment.StartTime = DateTime.Today.AddHours(15);
                appointment.EndTime = DateTime.Today.AddHours(16);
                appointment.Location = "Room 123";
                appointment.ReminderIsSet = true;
                appointment.ReminderMinutesBeforeStart = 30;
                appointment.RequiredAttendees.Add(new Attendee("John@mydomain.com"));
                appointment.OptionalAttendees.Add(new Attendee("Administrator@mydomain.com"));

                ItemId itemId = service.SendMeetingRequest(appointment);
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

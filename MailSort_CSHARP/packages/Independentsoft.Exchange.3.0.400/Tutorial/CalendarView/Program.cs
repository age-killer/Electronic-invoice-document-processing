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
                CalendarView view = new CalendarView(DateTime.Today, DateTime.Today.AddMonths(1));

                FindItemResponse response = service.FindItem(StandardFolder.Calendar, AppointmentPropertyPath.AllPropertyPaths, view);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Appointment)
                    {
                        Appointment appointment = (Appointment)response.Items[i];

                        Console.WriteLine("Subject = " + appointment.Subject);
                        Console.WriteLine("StartTime = " + appointment.StartTime);
                        Console.WriteLine("EndTime = " + appointment.EndTime);
                        Console.WriteLine("Body Preview = " + appointment.BodyPlainText);
                        Console.WriteLine("----------------------------------------------------------------");

                        if (appointment.InstanceType == InstanceType.Occurrence)
                        {
                            RecurringMasterItemId masterId = new RecurringMasterItemId(appointment.ItemId.Id, appointment.ItemId.ChangeKey);

                            Appointment master = service.GetAppointment(masterId);
                        }
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

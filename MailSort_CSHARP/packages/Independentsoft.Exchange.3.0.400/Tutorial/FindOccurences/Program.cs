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
                Exists restriction = new Exists(AppointmentPropertyPath.RecurrenceStart);

                FindItemResponse response = service.FindItem(StandardFolder.Calendar, AppointmentPropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Appointment)
                    {
                        Appointment master = (Appointment)response.Items[i];
                    
                        OccurrenceItemId occurrenceId = new OccurrenceItemId(master.ItemId.Id, master.ItemId.ChangeKey);
                    
                        //Get all occurences until exception occurs.
                        for(int j=1; j < 999; j++)
                        {
                            occurrenceId.Index = j;
                          
                            try
                            {
                                Appointment occurrence = service.GetAppointment(occurrenceId);
                           
                                Console.WriteLine("Subject = " + occurrence.Subject);
                                Console.WriteLine("StartTime = " + occurrence.StartTime);
                                Console.WriteLine("EndTime = " + occurrence.EndTime); 
                                Console.WriteLine("----------------------------------------------------------------");
                            }
                            catch (ServiceRequestException ex)
                            {
                                //ignore exception "Occurrence index is out of recurrence range"
                                break;
                            }
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

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
                //Daily recurrence. Every second day, next 20 days
                TaskRecurrence recurrence = new TaskRecurrence();
                recurrence.Pattern = new DailyRecurrencePattern(2);
                recurrence.Range = new NumberedRecurrenceRange(DateTime.Today, 20);

                Task task = new Task();
                task.Subject = "Every second day";
                task.Body = new Body("Body text");
                task.Owner = "My Name";
                task.StartDate = DateTime.Today.AddHours(10);
                task.DueDate = DateTime.Today.AddHours(20);
                task.Recurrence = recurrence;

                ItemId itemId = service.CreateItem(task);
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

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
                Task task = new Task();
                task.Subject = "Test";
                task.Body = new Body("Body text");
                task.Owner = "My Name";
                task.StartDate = DateTime.Today;
                task.DueDate = DateTime.Today.AddDays(3);
                task.ReminderIsSet = true;
                task.ReminderDueBy = DateTime.Today.AddDays(2);

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

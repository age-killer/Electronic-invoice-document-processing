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
                IsEqualTo restriction = new IsEqualTo(TaskPropertyPath.IsComplete, true);

                FindItemResponse response = service.FindItem(StandardFolder.Tasks, TaskPropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Task)
                    {
                        Task task = (Task)response.Items[i];

                        Console.WriteLine("Subject = " + task.Subject);
                        Console.WriteLine("StartDate = " + task.StartDate);
                        Console.WriteLine("DueDate = " + task.DueDate);
                        Console.WriteLine("Owner = " + task.Owner);
                        Console.WriteLine("Body Preview = " + task.BodyPlainText);
                        Console.WriteLine("----------------------------------------------------------------");
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

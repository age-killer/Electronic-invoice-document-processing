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
                IsEqualTo restriction = new IsEqualTo(TaskPropertyPath.ItemClass, ItemClass.TaskRequest);

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, TaskPropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Task)
                    {
                        ItemId taskRequestId = response.Items[i].ItemId;

                        Item item = service.GetItem(taskRequestId);
                        
                        ItemAttachment itemAttachment = (ItemAttachment)service.GetAttachment(item.Attachments[0].Id);
                        
                        Task task = (Task)itemAttachment.Item;

                        Console.WriteLine("Subject = " + task.Subject);
                        Console.WriteLine("StartDate = " + task.CommonStartDate);
                        Console.WriteLine("DueDate = " + task.DueDate);
                        Console.WriteLine("Owner = " + task.Owner);
                        Console.WriteLine("Body Preview = " + task.BodyPlainText);
                        Console.WriteLine("----------------------------------------------------------------");

                        //Accept task and create new task in the Tasks folder
                        ItemId taskId = service.CreateItem(task, StandardFolder.Tasks);

                        //Delete TaskRequest from the Inbox
                        service.DeleteItem(taskRequestId);
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

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
                IsEqualTo restriction = new IsEqualTo(MessagePropertyPath.IsRead, false);

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        Message message = (Message)response.Items[i];

                        Console.WriteLine("Subject = " + message.Subject);
                        Console.WriteLine("ReceivedTime = " + message.ReceivedTime);
                        Console.WriteLine("Body Preview = " + message.BodyPlainText);
                        Console.WriteLine("----------------------------------------------------------------");
                        
                        //If you want to get complete message with entire Body uncomment following line
                        //message = service.GetMessage(response.Items[i].ItemId);

                        //If you want to set message as read uncomment following lines
                        //Property readProperty = new Property(MessagePropertyPath.IsRead, true);
                        //ItemId itemId = service.UpdateItem(response.Items[i].ItemId, readProperty);
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

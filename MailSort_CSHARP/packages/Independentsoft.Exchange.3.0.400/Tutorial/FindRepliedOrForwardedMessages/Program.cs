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
                IsEqualTo restriction1 = new IsEqualTo(MessagePropertyPath.LastVerbExecuted, "102"); //LastVerbExecuted.ReplyToSender
                IsEqualTo restriction2 = new IsEqualTo(MessagePropertyPath.LastVerbExecuted, "103"); //LastVerbExecuted.ReplyToAll
                IsEqualTo restriction3 = new IsEqualTo(MessagePropertyPath.LastVerbExecuted, "104"); //LastVerbExecuted.Forward

                Or restriction4 = new Or(restriction1, restriction2, restriction3);

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, restriction4);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        Message message = (Message)response.Items[i];

                        Console.WriteLine("Subject = " + message.Subject);
                        Console.WriteLine("LastModifiedTime = " + message.LastModifiedTime);
                        Console.WriteLine("LastVerbExecuted = " + message.LastVerbExecuted);
                        Console.WriteLine("Body Preview = " + message.BodyPlainText);
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

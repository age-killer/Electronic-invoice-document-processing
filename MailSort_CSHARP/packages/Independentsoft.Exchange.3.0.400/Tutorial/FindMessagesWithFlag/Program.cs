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
                IsEqualTo restriction1 = new IsEqualTo(MessagePropertyPath.FlagStatus, "1"); //FlagStatus.Complete
                IsEqualTo restriction2 = new IsEqualTo(MessagePropertyPath.FlagStatus, "2"); //FlagStatus.Marked

                Or restriction3 = new Or(restriction1, restriction2);

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, restriction3);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        Message message = (Message)response.Items[i];

                        Console.WriteLine("Subject = " + message.Subject);
                        Console.WriteLine("FlagStatus = " + message.FlagStatus);
                        Console.WriteLine("FlagIcon = " + message.FlagIcon);
                        Console.WriteLine("FlagCompleteTime = " + message.FlagCompleteTime);
                        Console.WriteLine("FlagRequest = " + message.FlagRequest);
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

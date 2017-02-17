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
                IsEqualTo restriction = new IsEqualTo(MessagePropertyPath.Subject, "Test");

                FindItemResponse inboxResponse = service.FindItem(StandardFolder.Inbox, restriction);

                for (int i = 0; i < inboxResponse.Items.Count; i++)
                {
                    ItemId currentMessageId = inboxResponse.Items[i].ItemId;

                    ForwardItem forwardItem = new ForwardItem(currentMessageId);
                    forwardItem.NewBody = new Body("John, please see message below:");
                    forwardItem.ToRecipients.Add(new Mailbox("John@mydomain.com"));

                    ItemInfoResponse response = service.Forward(forwardItem);
                }
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

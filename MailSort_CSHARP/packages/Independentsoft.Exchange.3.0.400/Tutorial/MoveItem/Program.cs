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
                Message message = new Message();
                message.Subject = "Test";
                message.Body = new Body("Body text");
                message.ToRecipients.Add(new Mailbox("John@mydomain.com"));
                message.CcRecipients.Add(new Mailbox("Mark@mydomain.com"));

                ItemId itemId = service.CreateItem(message);

                ItemInfoResponse response = service.MoveItem(itemId, StandardFolder.Inbox);

                ItemId newItemId = response.Items[0].ItemId;
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

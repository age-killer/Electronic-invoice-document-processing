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

                //create message in the Drafts folder
                ItemId itemId1 = service.CreateItem(message);

                Mailbox john = new Mailbox("John@mydomain.com", "John Smith");
                Property toProperty = new Property(MessagePropertyPath.ToRecipients, john);
                
                //Update message.
                ItemId itemId2 = service.UpdateItem(itemId1, toProperty);

                //Send message
                ItemInfoResponse response = service.Send(itemId2);

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

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
                message.From = new Mailbox("Mark@mydomain.com");
                message.Subject = "Test message";
                message.Body = new Body("Body text");
                message.ToRecipients.Add(new Mailbox("John@mydomain.com"));
                message.ToRecipients.Add(new Mailbox("Peter@mydomain.com"));

                ItemInfoResponse response = service.Send(message);
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

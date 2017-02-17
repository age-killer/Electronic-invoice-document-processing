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
                service.ExchangeImpersonation = new Identity("John@mydomain.com");

                Message message = new Message();
                message.Subject = "Test";
                message.Body = new Body("Body text");
                message.ToRecipients.Add(new Mailbox("Mark@mydomain.com"));

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

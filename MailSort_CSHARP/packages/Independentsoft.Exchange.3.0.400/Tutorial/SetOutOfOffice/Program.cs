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
                OutOfOffice oof = new OutOfOffice();
                oof.State = OutOfOfficeState.Scheduled;
                oof.Duration = new TimeDuration(DateTime.Today, DateTime.Today.AddDays(10));
                oof.InternalReply = new ReplyBody("I am out of office until " + DateTime.Today.AddDays(10));
                oof.ExternalReply = new ReplyBody("I am out of office until " + DateTime.Today.AddDays(10));

                Response response = service.SetOutOfOffice(oof, "John@mydomain.com");
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

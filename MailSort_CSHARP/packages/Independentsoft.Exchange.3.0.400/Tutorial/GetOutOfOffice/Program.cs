using System;
using System.Net;
using Independentsoft.Exchange;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "passsword");
            Service service = new Service("https://myserver/ews/Exchange.asmx", credential);

            try
            {
                OutOfOfficeResponse response = service.GetOutOfOffice("John@mydomain.com");

                OutOfOffice oof = response.OutOfOffice;

                Console.WriteLine("State = " + oof.State);

                if (oof.InternalReply != null)
                {
                    Console.WriteLine("InternalReply = " + oof.InternalReply.Message);
                }

                if (oof.ExternalReply != null)
                {
                    Console.WriteLine("ExternalReply = " + oof.ExternalReply.Message);
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

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

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1;

            try
            {
                FindItemResponse response = service.FindItem(StandardFolder.Inbox, new ItemShape(ShapeType.Id));

                for (int i = 0; i < response.Items.Count; i++)
                {
                    AlternateId sourceId = new AlternateId(IdFormat.EwsId, response.Items[i].ItemId.Id, "John@domain.com");
  
                    ConvertIdResponse convertIdResponse = service.ConvertId(sourceId, IdFormat.OwaId);

                    string owaId = convertIdResponse.AlternateId.Id;

                    Console.WriteLine("Outlook Web App ItemId:" + owaId);                    
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

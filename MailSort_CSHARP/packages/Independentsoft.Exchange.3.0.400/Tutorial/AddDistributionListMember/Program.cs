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
                DistributionList list = service.GetDistributionList("AAMkADZkZDU1ZTQyLTkZ8lTi+aWAAAAAAAZAABRQo1uip9ARZANZ8lTi+aWAACof8EsAAA=");

                ItemId itemId = list.ItemId;

                Property memebersProperty = new Property(DistributionListPropertyPath.Members, new DistributionListMember(new Mailbox("john@mydomain.com")));

                ItemChange itemChange = new ItemChange(itemId);
                itemChange.PropertiesToAppend.Add(memebersProperty);

                itemId = service.UpdateItem(itemChange);  
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

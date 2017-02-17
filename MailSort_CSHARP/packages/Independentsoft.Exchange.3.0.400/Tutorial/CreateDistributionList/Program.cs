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
                DistributionListMember member1 = new DistributionListMember(new Mailbox("John@mydomain.com"));
                DistributionListMember member2 = new DistributionListMember(new Mailbox("Mark@mydomain.com"));

                DistributionList list = new DistributionList();
                list.DisplayName = "Test";
                list.Subject = "Test";
                list.Body = new Body("Body text");
                list.Members.Add(member1);
                list.Members.Add(member2);

                ItemId itemId = service.CreateItem(list);
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

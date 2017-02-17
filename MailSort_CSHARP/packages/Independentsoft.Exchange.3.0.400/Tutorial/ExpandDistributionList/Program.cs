using System;
using System.Net;
using System.Collections.Generic;
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
                IsEqualTo restriction = new IsEqualTo(ContactPropertyPath.ItemClass, ItemClass.DistributionList);

                FindItemResponse response = service.FindItem(StandardFolder.Contacts, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    ItemId itemId = response.Items[i].ItemId;

                    ExpandDistributionListResponse listResponse = service.ExpandDistributionList(itemId);
                    IList<Mailbox> mailboxes = listResponse.Mailboxes;

                    for (int j = 0; j < mailboxes.Count; j++)
                    {
                        Console.WriteLine("Member = " + mailboxes[j].EmailAddress);
                    }

                    Console.WriteLine("---------------------------------------------------------------");
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

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
                IsEqualTo restriction = new IsEqualTo(ContactPropertyPath.Email1Address, "john@mydomain.com");

                FindItemResponse response = service.FindItem(StandardFolder.Contacts, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Contact)
                    {
                        ItemId itemId = response.Items[i].ItemId;

                        Property businessPhoneProperty = new Property(ContactPropertyPath.BusinessPhone, "555-666-777");

                        itemId = service.UpdateItem(itemId, businessPhoneProperty);
                    }
                }
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


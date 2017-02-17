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
                IsEqualTo restriction = new IsEqualTo(ContactPropertyPath.GivenName, "Michael");

                FindItemResponse response = service.FindItem(StandardFolder.Contacts, ContactPropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Contact)
                    {
                        Contact contact = (Contact)response.Items[i];

                        Console.WriteLine("GivenName = " + contact.GivenName);
                        Console.WriteLine("Surname = " + contact.Surname);
                        Console.WriteLine("CompanyName = " + contact.CompanyName);
                        Console.WriteLine("BusinessAddress = " + contact.BusinessAddress);
                        Console.WriteLine("BusinessPhone = " + contact.BusinessPhone);
                        Console.WriteLine("Email1DisplayName = " + contact.Email1DisplayName);
                        Console.WriteLine("Email1Address = " + contact.Email1Address);
                        Console.WriteLine("----------------------------------------------------------------");
                    }
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

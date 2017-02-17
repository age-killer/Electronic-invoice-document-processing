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
                PropertyName myPropertyName = new PropertyName("myfield1", StandardPropertySet.PublicStrings, MapiPropertyType.String);
                Exists restriction = new Exists(myPropertyName);

                IList<PropertyPath> propertyPaths = new List<PropertyPath>();
                propertyPaths.Add(ContactPropertyPath.GivenName);
                propertyPaths.Add(ContactPropertyPath.Surname);
                propertyPaths.Add(ContactPropertyPath.Email1DisplayName);
                propertyPaths.Add(ContactPropertyPath.Email1Address);
                propertyPaths.Add(myPropertyName);

                FindItemResponse response = service.FindItem(StandardFolder.Contacts, propertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Contact)
                    {
                        Contact contact = (Contact)response.Items[i];

                        Console.WriteLine("GivenName = " + contact.GivenName);
                        Console.WriteLine("Surname = " + contact.Surname);
                        Console.WriteLine("CompanyName = " + contact.CompanyName);
                        Console.WriteLine("Email1DisplayName = " + contact.Email1DisplayName);
                        Console.WriteLine("Email1Address = " + contact.Email1Address);
                        Console.WriteLine("Myfield1 = " + contact.ExtendedProperties[myPropertyName].Value);
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

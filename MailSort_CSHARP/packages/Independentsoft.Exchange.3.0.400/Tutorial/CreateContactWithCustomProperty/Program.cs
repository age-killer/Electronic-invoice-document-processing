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
                PropertyName myPropertyName = new PropertyName("myfield1", StandardPropertySet.PublicStrings, MapiPropertyType.String);
                ExtendedProperty myProperty = new ExtendedProperty(myPropertyName, "test");

                Contact contact = new Contact();
                contact.GivenName = "John";
                contact.Surname = "Smith";
                contact.FileAsMapping = FileAsMapping.LastSpaceFirst;
                contact.CompanyName = "Independentsoft";
                contact.BusinessPhone = "555-666-777";
                contact.Email1Address = "john@independentsoft.de";
                contact.Email1DisplayName = "John Smith";
                contact.Email1DisplayAs = "John Smith";
                contact.Email1Type = "SMTP";
                contact.ExtendedProperties.Add(myProperty);

                ItemId itemId = service.CreateItem(contact);
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

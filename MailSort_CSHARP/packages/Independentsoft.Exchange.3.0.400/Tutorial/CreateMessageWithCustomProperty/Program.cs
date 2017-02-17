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
                PropertyName myPropertyName = new PropertyName("MyCustomProperty", StandardPropertySet.PublicStrings, MapiPropertyType.String);
                ExtendedProperty myProperty = new ExtendedProperty(myPropertyName, "value1");

                Message message = new Message();
                message.Subject = "Test";
                message.Body = new Body("Body text");
                message.ToRecipients.Add(new Mailbox("John@mydomain.com"));
                message.CcRecipients.Add(new Mailbox("Mark@mydomain.com"));
                message.ExtendedProperties.Add(myProperty);

                ItemId itemId = service.CreateItem(message);
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

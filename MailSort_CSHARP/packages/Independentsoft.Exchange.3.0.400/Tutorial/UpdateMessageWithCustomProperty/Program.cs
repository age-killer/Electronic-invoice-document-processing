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
                Contains restriction = new Contains(MessagePropertyPath.Subject, "test", ContainmentMode.Prefixed, ContainmentComparison.IgnoreCase);

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        PropertyName myPropertyName = new PropertyName("MyCustomProperty", StandardPropertySet.PublicStrings, MapiPropertyType.String);
                        ExtendedProperty myProperty = new ExtendedProperty(myPropertyName, "value1");

                        ItemId newItemId = service.UpdateItem(response.Items[i].ItemId, myProperty);
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

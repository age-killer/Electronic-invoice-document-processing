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
                IsEqualTo restriction = new IsEqualTo(JournalPropertyPath.ItemClass, ItemClass.Journal);

                FindItemResponse response = service.FindItem(StandardFolder.Journal, JournalPropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    Journal journal = (Journal)response.Items[i];

                    Console.WriteLine("Subject = " + journal.Subject);
                    Console.WriteLine("Type = " + journal.Type);
                    Console.WriteLine("TypeDescription = " + journal.TypeDescription);
                    Console.WriteLine("StartTime = " + journal.StartTime);
                    Console.WriteLine("EndTime = " + journal.EndTime);
                    Console.WriteLine("Body Preview = " + journal.BodyPlainText);
                    Console.WriteLine("----------------------------------------------------------------");
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

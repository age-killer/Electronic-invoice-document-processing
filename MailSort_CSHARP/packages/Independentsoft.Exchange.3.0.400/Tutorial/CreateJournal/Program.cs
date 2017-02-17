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
                Journal journal = new Journal();
                journal.Subject = "Test";
                journal.Body = new Body("Body text");
                journal.Type = "Phone call";
                journal.TypeDescription = "Phone call";
                journal.Companies.Add("Independentsoft");
                journal.StartTime = DateTime.Today.AddHours(15);
                journal.EndTime = DateTime.Today.AddHours(16);

                ItemId itemId = service.CreateItem(journal, StandardFolder.Journal);
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

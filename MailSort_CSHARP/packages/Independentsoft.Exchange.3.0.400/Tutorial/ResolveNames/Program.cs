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
                ResolveNamesResponse response = service.ResolveNames("John");

                for (int i = 0; i < response.Resolutions.Count; i++)
                {
                    Resolution resolution = response.Resolutions[i];

                    if (resolution != null && resolution.Mailbox != null)
                    {
                        Mailbox mailbox = resolution.Mailbox;

                        Console.WriteLine("Name = " + mailbox.Name);
                        Console.WriteLine("EmailAddress = " + mailbox.EmailAddress);
                    }

                    if (resolution != null && resolution.Contact != null)
                    {
                        Contact contact = resolution.Contact;

                        Console.WriteLine("GivenName = " + contact.GivenName);
                        Console.WriteLine("EmailAddress = " + contact.Email1Address);
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

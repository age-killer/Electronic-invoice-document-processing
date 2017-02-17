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

            service.RequestServerVersion = RequestServerVersion.Exchange2013;

            try
            {
                IsEqualTo restriction = new IsEqualTo(PersonaPropertyPath.Surname, "Smith");

                FindPeopleResponse response = service.FindPeople(StandardFolder.Contacts, restriction);

                foreach (Persona persona in response.Personas)
                {
                    Console.WriteLine("PersonaId:" + persona.PersonaId);
                    Console.WriteLine("PersonaType:" + persona.PersonaType);
                    Console.WriteLine("DisplayName:" + persona.DisplayName);
                    Console.WriteLine("PhoneNumber:" + persona.PhoneNumber);

                    if (persona.EmailAddress != null)
                    {
                        Console.WriteLine("EmailAddress:" + persona.EmailAddress.EmailAddress);
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

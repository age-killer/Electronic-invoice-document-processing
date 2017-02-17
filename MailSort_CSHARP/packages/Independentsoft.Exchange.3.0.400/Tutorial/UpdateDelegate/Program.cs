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
                Mailbox mailbox = new Mailbox("John@mydomain.com");

                DelegateUser delegateUser = new DelegateUser("Mark@mydomain.com");
                delegateUser.ContactsFolderPermissionLevel = DelegateFolderPermissionLevel.Author;

                DelegateUserResponse response = service.UpdateDelegate(mailbox, delegateUser);
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

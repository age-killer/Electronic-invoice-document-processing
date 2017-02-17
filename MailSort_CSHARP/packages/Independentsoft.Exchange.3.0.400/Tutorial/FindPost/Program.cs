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
                IsGreaterThanOrEqualTo restriction = new IsGreaterThanOrEqualTo(PostPropertyPath.PostedTime, DateTime.Today.AddMonths(-1));

                FindItemResponse response = service.FindItem(StandardFolder.Drafts, PostPropertyPath.AllPropertyPaths, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Post)
                    {
                        Post post = (Post)response.Items[i];

                        Console.WriteLine("Subject = " + post.Subject);
                        Console.WriteLine("PostedTime = " + post.PostedTime);
                        Console.WriteLine("Body Preview = " + post.BodyPlainText);
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

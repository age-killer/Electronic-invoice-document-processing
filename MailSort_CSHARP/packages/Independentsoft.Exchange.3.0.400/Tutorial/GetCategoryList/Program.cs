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

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1;

            try
            {
                CategoryList list = service.GetCategoryList();

                foreach (Category category in list.Categories)
                {
                    Console.WriteLine(category.Name);
                    Console.WriteLine(category.Color);
                    Console.WriteLine("--------------------------------------------");
                }

                //if you want to delete category from the list
                for (int i = list.Categories.Count - 1; i >= 0; i--)
                {
                    if (list.Categories[i].Name == "MyCategory")
                    {
                        list.Categories.RemoveAt(i);
                    }
                }

                //save changes after delete
                service.UpdateCategoryList(list);
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

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

                Category myCategory = new Category("Steel Category", CategoryColor.DarkSteel, KeyboardShortcut.CtrlF7);
                myCategory.Guid = "{fac7e643-c161-4f81-b394-898943335ad4}";

                list.Categories.Add(myCategory);

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

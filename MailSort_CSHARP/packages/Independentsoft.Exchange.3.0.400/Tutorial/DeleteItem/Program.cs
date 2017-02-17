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

            try
            {
                //Delete Draft items older then 1 month

                IsLessThan restriction = new IsLessThan(ItemPropertyPath.CreatedTime, DateTime.Today.AddMonths(-1));

                FindItemResponse draftsItems = service.FindItem(StandardFolder.Drafts, restriction);

                for (int i = 0; i < draftsItems.Items.Count; i++)
                {
                    Response response = service.DeleteItem(draftsItems.Items[i].ItemId, DeleteType.MoveToDeletedItems);
                }

                //Empty DeletedItems folder
                
                ItemShape itemShape = new ItemShape(ShapeType.Id);
                FindItemResponse deletedItems = service.FindItem(StandardFolder.DeletedItems, itemShape);

                IList<ItemId> itemIds = new List<ItemId>();

                for (int i = 0; i < deletedItems.Items.Count; i++)
                {
                    itemIds.Add(deletedItems.Items[i].ItemId);
                }

                IList<Response> responses = service.DeleteItem(itemIds, DeleteType.HardDelete);
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

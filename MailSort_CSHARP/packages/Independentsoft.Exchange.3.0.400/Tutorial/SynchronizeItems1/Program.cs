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
                //Initial synscronization. It will retrieve all existing items from the specified folder 
                SyncItemsResponse response = service.SyncItems(StandardFolder.Inbox);

                //Keep synchronization state in order to use it in next synchronizations
                string state = response.State;

                IList<Item> allExistingItems = response.CreatedItems;

                Console.WriteLine("Wait until new messages arrive or create/mark read/delete messages in the Inbox folder and press ENTER");
                Console.Read();

                //Synchronize items again but now include the state parameter in order to find changes from last synchronization
                SyncItemsResponse newResponse = service.SyncItems(StandardFolder.Inbox, state);

                IList<Item> newItems = newResponse.CreatedItems;
                IList<ItemId> deletedItems = newResponse.DeletedItems;
                IList<Item> updatedItems = newResponse.UpdatedItems;
                IList<ReadFlagChange> readFlagChangedItems = newResponse.ReadFlagChangedItems;

                //Display new items
                for (int i = 0; i < newItems.Count; i++)
                {
                    Console.WriteLine("New = " + newItems[i].Subject);
                    Console.WriteLine("New ItemId= " + newItems[i].ItemId);
                }

                //Display deleted items
                for (int i = 0; i < deletedItems.Count; i++)
                {
                    Console.WriteLine("Deleted ItemId = " + deletedItems[i]);
                }

                //Display updated items
                for (int i = 0; i < updatedItems.Count; i++)
                {
                    Console.WriteLine("Updated = " + updatedItems[i].Subject);
                    Console.WriteLine("Updated ItemId = " + updatedItems[i].ItemId);
                }

                //Display items marked read/unread
                for (int i = 0; i < readFlagChangedItems.Count; i++)
                {
                    Console.WriteLine("Marked read/unread = " + readFlagChangedItems[i].ItemId);
                    Console.WriteLine("IsRead = " + readFlagChangedItems[i].IsRead);
                }

                Console.Read();
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

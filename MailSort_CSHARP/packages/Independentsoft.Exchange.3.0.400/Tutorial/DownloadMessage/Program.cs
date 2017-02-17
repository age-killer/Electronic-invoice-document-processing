using System;
using System.IO;
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
                ItemShape itemShape = new ItemShape(ShapeType.Id);
                FindItemResponse inboxItems = service.FindItem(StandardFolder.Inbox, itemShape);

                for (int i = 0; i < inboxItems.Items.Count; i++)
                {
                    ItemId currentItemId = inboxItems.Items[i].ItemId;

                    Message message = service.GetMessage(currentItemId);
                    
                    string mimeContent = message.MimeContent.Text;
                    byte[] buffer = System.Text.Encoding.UTF8.GetBytes(mimeContent);

                    string fileName = "c:\\test\\message" + i + ".eml";

                    FileStream file = new FileStream(fileName, FileMode.Create);
                    
                    using(file)
                    {
                        file.Write(buffer, 0, buffer.Length);
                    }
                }
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

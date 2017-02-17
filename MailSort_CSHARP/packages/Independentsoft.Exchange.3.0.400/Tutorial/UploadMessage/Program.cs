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
                FileInfo fileInfo = new FileInfo("c:\\test\\message.eml");
                byte[] buffer = new byte[fileInfo.Length];

                FileStream file = new FileStream("c:\\test\\message.eml", FileMode.Open);

                using (file)
                {
                    file.Read(buffer, 0, buffer.Length);
                }

                string mimeText = System.Text.Encoding.UTF8.GetString(buffer);
                MimeContent mimeContent = new MimeContent(mimeText);

                Message message = new Message(mimeContent);
  
                ItemId itemId = service.CreateItem(message, StandardFolder.Drafts);
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

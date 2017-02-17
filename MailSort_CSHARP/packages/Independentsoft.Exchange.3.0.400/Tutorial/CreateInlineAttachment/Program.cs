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
                Message message = new Message();
                message.Subject = "Inline";
                message.HideAttachments = true;
                message.Body = new Body("<html><body>Here is a message with an inline attachment:<br><img widht=\"640\" height=\"480\" id=\"Picture1\" src=\"cid:picture.jpg@123456\"></img></body></html>", BodyType.Html);

                ItemId itemId = service.CreateItem(message);

                FileAttachment fileAttachment = new FileAttachment("c:\\test\\picture.jpg");
                fileAttachment.ContentId = "picture.jpg@123456";

                AttachmentId attachmentId = service.CreateAttachment(fileAttachment, itemId);
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

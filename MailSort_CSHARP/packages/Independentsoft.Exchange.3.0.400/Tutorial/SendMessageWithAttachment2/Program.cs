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
                //Create message in the Drafts folder
                Message message = new Message();
                message.Subject = "Test";
                message.Body = new Body("Body text");
                message.ToRecipients.Add(new Mailbox("John@mydomain.com"));

                ItemId messageId = service.CreateItem(message);

                //Attach message
                Message attachedMessage = new Message();
                attachedMessage.Subject = "MyAttachment";
                attachedMessage.Body = new Body("Attached message body text.");

                ItemAttachment attachment = new ItemAttachment("MyAttachment", attachedMessage);
                AttachmentId attachmentId = service.CreateAttachment(attachment, messageId);

                //Update messageId
                messageId.ChangeKey = attachmentId.RootItemChangeKey;

                //Send message
                ItemInfoResponse response = service.Send(messageId);
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

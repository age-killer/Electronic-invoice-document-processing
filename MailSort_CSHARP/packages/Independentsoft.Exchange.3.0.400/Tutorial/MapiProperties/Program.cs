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
                IList<PropertyPath> propertyPaths = new List<PropertyPath>();
                propertyPaths.Add(MessagePropertyPath.Subject);
                propertyPaths.Add(MessagePropertyPath.ReceivedTime);
                propertyPaths.Add(MessagePropertyPath.SentTime);
                propertyPaths.Add(MessagePropertyPath.HasAttachments);
                propertyPaths.Add(MessagePropertyPath.IsRead);
                propertyPaths.Add(MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS); 
                propertyPaths.Add(MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS); 
                propertyPaths.Add(MapiPropertyTag.PR_BODY);
                propertyPaths.Add(MapiPropertyTag.PR_HTML);

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, propertyPaths);

                for (int i = 0; i < response.Items.Count; i++)
               {
                    if (response.Items[i] is Message)
                    {
                        Message message = (Message)response.Items[i];

                        Console.WriteLine("Subject = " + message.Subject);
                        Console.WriteLine("ReceivedTime = " + message.ReceivedTime);
                        Console.WriteLine("SentTime = " + message.SentTime);
                        Console.WriteLine("HasAttachments = " + message.HasAttachments);
                        Console.WriteLine("IsRead = " + message.IsRead);

                        if(message.ExtendedProperties[MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS] != null)
                        {
                            Console.WriteLine("From email = " + message.ExtendedProperties[MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS].Value);
                        }

                        if(message.ExtendedProperties[MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS] != null)
                        {
                            Console.WriteLine("To = " + message.ExtendedProperties[MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS].Value);
                        }

                        if(message.ExtendedProperties[MapiPropertyTag.PR_BODY] != null)
                        {
                            Console.WriteLine("Plain body = " + message.ExtendedProperties[MapiPropertyTag.PR_BODY].Value);
                        }

                        if(message.ExtendedProperties[MapiPropertyTag.PR_HTML] != null)
                        {
                            byte[] base64 = Convert.FromBase64String(message.ExtendedProperties[MapiPropertyTag.PR_HTML].Value);
                            string htmlBody = System.Text.Encoding.UTF8.GetString(base64);

                            Console.WriteLine("Html body = " + htmlBody);
                        }

                        Console.WriteLine("--------------------------------------------------------");
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

using System;
using System.IO;
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
                Contains restriction = new Contains(MessagePropertyPath.Attachments, "test", ContainmentMode.Prefixed, ContainmentComparison.IgnoreCase);

                ItemShape itemShape = new ItemShape(ShapeType.Id);

                FindItemResponse response = service.FindItem(StandardFolder.Inbox, itemShape, restriction);

                for (int i = 0; i < response.Items.Count; i++)
                {
                    if (response.Items[i] is Message)
                    {
                        ItemId itemId = response.Items[i].ItemId;

                        Message message = service.GetMessage(itemId);

                        IList<AttachmentInfo> attachmentsInfo = message.Attachments;

                        for (int j = 0; j < attachmentsInfo.Count; j++)
                        {
                            Attachment attachment = service.GetAttachment(attachmentsInfo[j].Id);

                            if (attachment is FileAttachment)
                            {
                                FileAttachment fileAttachment = (FileAttachment)attachment;

                                byte[] buffer = fileAttachment.Content;
                                string fileName = fileAttachment.Name;

                                FileStream fileStream = new FileStream("c:\\test\\" + fileName, FileMode.Create);

                                using (fileStream)
                                {
                                    fileStream.Write(buffer, 0, buffer.Length);
                                }
                            }
                            else if (attachment is ItemAttachment)
                            {
                                ItemAttachment itemAttachment = (ItemAttachment)attachment;

                                string name = itemAttachment.Name;
                                Item attachedItem = itemAttachment.Item;
                            }
                        }
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
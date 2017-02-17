using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using Independentsoft.Exchange;

namespace Sample
{
    class Program
    {
        private static Service service = null;

        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            service = new Service("https://myserver/ews/Exchange.asmx", credential);

            try
            {
                //send to my client computer on port 1234
                //the client computer will listen on port 1234 (see method Listen())
                PushSubscription subscription = new PushSubscription(StandardFolder.Inbox, EventType.NewMail, "http://myclient:1234");
                SubscribeResponse subscribeResponse = service.Subscribe(subscription);

                Listen();
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

        public static void Listen()
        {
            Console.WriteLine("Listen ...");

            try
            {
                IPAddress localAddress = GetIPAddress();
                IPEndPoint localIpEndPoint = new IPEndPoint(localAddress, 1234);

                Socket tcpListenerSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                tcpListenerSocket.Bind(localIpEndPoint);
                tcpListenerSocket.Listen(10);

                while (true)
                {
                    Socket tcpAcceptedSocket = tcpListenerSocket.Accept();

                    MemoryStream memoryStream = new MemoryStream();
                    bool isComplete = false;

                    while (!isComplete)
                    {
                        byte[] buffer = new byte[16384];

                        int bytesReceived = tcpAcceptedSocket.Receive(buffer);
                        memoryStream.Write(buffer, 0, bytesReceived);

                        isComplete = IsCompleteRequestReceived(memoryStream.ToArray());
                    }

                    //Send OK response
                    string okResponse = service.SendNotificationResult(SubscriptionStatus.OK); //to unsubscribe use SubscriptionStatus.Unsubscribe
                    byte[] responseBuffer = System.Text.Encoding.UTF8.GetBytes(okResponse);

                    tcpAcceptedSocket.Send(responseBuffer);
                    tcpAcceptedSocket.Shutdown(SocketShutdown.Both);
                    tcpAcceptedSocket.Close();

                    //Process received notification
                    byte[] receivedBuffer = memoryStream.ToArray();

                    SendNotificationResponse response = service.ParseSendNotificationResponse(receivedBuffer);

                    if (response != null)
                    {
                        Notification notification = response.Notification;

                        for (int i = 0; i < notification.Events.Count; i++)
                        {
                            Console.WriteLine("SubscriptionId: " + notification.SubscriptionId);

                            if (notification.Events[i] is NewMailEvent)
                            {
                                NewMailEvent newMailEvent = (NewMailEvent)notification.Events[i];

                                ItemId itemId = (ItemId)newMailEvent.Id;

                                Message message = service.GetMessage(itemId);

                                Console.WriteLine("ItemId: " + message.ItemId);
                                Console.WriteLine("Subject: " + message.Subject);
                                Console.WriteLine("ReceivedTime: " + message.ReceivedTime);
                            }
                            else if (notification.Events[i] is StatusEvent)
                            {
                                StatusEvent statusEvent = (StatusEvent)notification.Events[i];

                                Console.WriteLine("I'm alive. Watermark: " + statusEvent.Watermark);
                            }
                        }
                    }
                }
            }
            catch (SocketException se)
            {
                Console.WriteLine(se.Message);
                Console.Read();
            }
        }

        private static IPAddress GetIPAddress()
        {
            IPHostEntry localhost = Dns.GetHostEntry(Dns.GetHostName());

            IPAddress localAddress = localhost.AddressList[0];

            if (localhost.AddressList.Length > 1)
            {
                for (int i = 0; i < localhost.AddressList.Length; i++)
                {
                    if (localhost.AddressList[i] != null && localhost.AddressList[i].ToString().IndexOf(":") == -1)
                    {
                        localAddress = localhost.AddressList[i];
                        break;
                    }
                }
            }

            return localAddress;
        }

        private static bool IsCompleteRequestReceived(byte[] buffer)
        {
            string request = System.Text.Encoding.UTF8.GetString(buffer);

            int bodyIndex = request.IndexOf("\r\n\r\n");

            if (bodyIndex == -1)
            {
                return false;
            }

            string header = request.Substring(0, bodyIndex);
            string body = request.Substring(bodyIndex + 4);

            StringReader stringReader = new StringReader(header);

            int contentLength = 0;
            string headerLine;

            while ((headerLine = stringReader.ReadLine()) != null)
            {
                headerLine = headerLine.Trim();

                if (headerLine.StartsWith("Content-Length:"))
                {
                    string contentLengthValue = headerLine.Substring(15);

                    contentLengthValue = contentLengthValue.Trim();

                    if (contentLengthValue != null && contentLengthValue.Length > 0)
                    {
                        contentLength = Int32.Parse(contentLengthValue);
                        break;
                    }
                }
            }

            if (contentLength == body.Length)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
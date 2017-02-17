Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1

        Private Shared service As Service = Nothing

        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try
                'send to my client computer on port 1234
                'the client computer will listen on port 1234 (see method Listen())
                Dim subscription As New PushSubscription(StandardFolder.Inbox, EventType.NewMail, "http://myclient:1234")
                Dim subscribeResponse As SubscribeResponse = service.Subscribe(subscription)

                Listen()

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " & ex.Message)
                Console.WriteLine("Error: " & ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " & ex.Message)
                Console.Read()
            End Try
        End Sub

        Public Shared Sub Listen()

            Console.WriteLine("Listen ...")

            Try
                Dim localAddress As IPAddress = GetIPAddress()
                Dim localIpEndPoint As New IPEndPoint(localAddress, 1234)

                Dim tcpListenerSocket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                tcpListenerSocket.Bind(localIpEndPoint)
                tcpListenerSocket.Listen(10)

                While True
                    Dim tcpAcceptedSocket As Socket = tcpListenerSocket.Accept()

                    Dim memoryStream As New MemoryStream()
                    Dim isComplete As Boolean = False

                    While Not isComplete
                        Dim buffer As Byte() = New Byte(16383) {}

                        Dim bytesReceived As Integer = tcpAcceptedSocket.Receive(buffer)
                        memoryStream.Write(buffer, 0, bytesReceived)

                        isComplete = IsCompleteRequestReceived(memoryStream.ToArray())
                    End While

                    'Send OK response
                    Dim okResponse As String = service.SendNotificationResult(SubscriptionStatus.OK) 'to unsubscribe use SubscriptionStatus.Unsubscribe
                    Dim responseBuffer As Byte() = System.Text.Encoding.UTF8.GetBytes(okResponse)

                    tcpAcceptedSocket.Send(responseBuffer)
                    tcpAcceptedSocket.Shutdown(SocketShutdown.Both)
                    tcpAcceptedSocket.Close()

                    'Process received notification
                    Dim receivedBuffer As Byte() = memoryStream.ToArray()

                    Dim response As SendNotificationResponse = service.ParseSendNotificationResponse(receivedBuffer)

                    If response IsNot Nothing Then
                        Dim notification As Notification = response.Notification

                        For i As Integer = 0 To notification.Events.Count - 1

                            Console.WriteLine("SubscriptionId: " & notification.SubscriptionId)

                            If TypeOf notification.Events(i) Is NewMailEvent Then

                                Dim newMailEvent As NewMailEvent = DirectCast(notification.Events(i), NewMailEvent)

                                Dim itemId As ItemId = DirectCast(newMailEvent.Id, ItemId)

                                Dim message As Message = service.GetMessage(itemId)

                                Console.WriteLine("ItemId: " & message.ItemId.ToString())
                                Console.WriteLine("Subject: " & message.Subject)
                                Console.WriteLine("ReceivedTime: " & message.ReceivedTime)

                            ElseIf TypeOf notification.Events(i) Is StatusEvent Then

                                Dim statusEvent As StatusEvent = DirectCast(notification.Events(i), StatusEvent)

                                Console.WriteLine("I'm alive. Watermark: " & statusEvent.Watermark)
                            End If
                        Next
                    End If
                End While
            Catch se As SocketException
                Console.WriteLine(se.Message)
                Console.Read()
            End Try
        End Sub

        Private Shared Function GetIPAddress() As IPAddress
            Dim localhost As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName())

            Dim localAddress As IPAddress = localhost.AddressList(0)

            If localhost.AddressList.Length > 1 Then
                For i As Integer = 0 To localhost.AddressList.Length - 1
                    If localhost.AddressList(i) IsNot Nothing AndAlso localhost.AddressList(i).ToString().IndexOf(":") = -1 Then
                        localAddress = localhost.AddressList(i)
                        Exit For
                    End If
                Next
            End If

            Return localAddress
        End Function

        Private Shared Function IsCompleteRequestReceived(ByVal buffer As Byte()) As Boolean
            Dim request As String = System.Text.Encoding.UTF8.GetString(buffer)

            Dim bodyIndex As Integer = request.IndexOf(vbCr & vbLf & vbCr & vbLf)

            If bodyIndex = -1 Then
                Return False
            End If

            Dim header As String = request.Substring(0, bodyIndex)
            Dim body As String = request.Substring(bodyIndex + 4)

            Dim stringReader As New StringReader(header)

            Dim contentLength As Integer = 0
            Dim headerLine As String = stringReader.ReadLine()

            While (headerLine) IsNot Nothing

                headerLine = headerLine.Trim()

                If headerLine.StartsWith("Content-Length:") Then

                    Dim contentLengthValue As String = headerLine.Substring(15)

                    contentLengthValue = contentLengthValue.Trim()

                    If contentLengthValue IsNot Nothing AndAlso contentLengthValue.Length > 0 Then
                        contentLength = Int32.Parse(contentLengthValue)
                        Exit While
                    End If
                End If

                headerLine = stringReader.ReadLine()

            End While

            If contentLength = body.Length Then
                Return True
            Else
                Return False
            End If
        End Function
    End Class
End Namespace

Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim newWatermark As String = Nothing

                Dim subscription As New PullSubscription(StandardFolder.Inbox, EventType.NewMail)

                'initail subscribe 
                Dim subscribeResponse As SubscribeResponse = service.Subscribe(subscription)

                While True

                    'wait 60 seconds 
                    System.Threading.Thread.CurrentThread.Join(60000)

                    Dim eventsResponse As GetEventsResponse = service.GetEvents(subscribeResponse)

                    Dim notification As Notification = eventsResponse.Notification

                    For i As Integer = 0 To notification.Events.Count - 1

                        newWatermark = notification.Events(i).Watermark

                        If TypeOf notification.Events(i) Is NewMailEvent Then

                            Dim newMailEvent As NewMailEvent = DirectCast(notification.Events(i), NewMailEvent)

                            Dim itemId As ItemId = DirectCast(newMailEvent.Id, ItemId)

                            Dim message As Message = service.GetMessage(itemId)

                            Console.WriteLine(message.Subject)
                            Console.WriteLine(message.ReceivedTime)
                        End If
                    Next

                    'resubscribe with new watermark 
                    subscription.Watermark = newWatermark
                    subscribeResponse = service.Subscribe(subscription)

                End While

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " + ex.Message)
                Console.WriteLine("Error: " + ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " + ex.Message)
                Console.Read()
            End Try

        End Sub
    End Class
End Namespace
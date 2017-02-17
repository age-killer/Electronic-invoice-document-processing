Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim newWatermark As String = Nothing

                Dim eventTypes As New List(Of EventType)()
                eventTypes.Add(EventType.Created)
                eventTypes.Add(EventType.Modified)

                Dim subscription As New PullSubscription(StandardFolder.Calendar, eventTypes)

                'initail subscribe 
                Dim subscribeResponse As SubscribeResponse = service.Subscribe(subscription)

                While True

                    'wait 60 seconds 
                    System.Threading.Thread.CurrentThread.Join(60000)

                    Dim eventsResponse As GetEventsResponse = service.GetEvents(subscribeResponse)

                    Dim notification As Notification = eventsResponse.Notification

                    For i As Integer = 0 To notification.Events.Count - 1

                        newWatermark = notification.Events(i).Watermark

                        If TypeOf notification.Events(i) Is CreatedEvent Then

                            Dim createdEvent As CreatedEvent = DirectCast(notification.Events(i), CreatedEvent)

                            If TypeOf createdEvent.Id Is ItemId Then

                                Dim itemId As ItemId = DirectCast(createdEvent.Id, ItemId)

                                Console.WriteLine("Created item = " & itemId.ToString())

                            ElseIf TypeOf createdEvent.Id Is FolderId Then
                                Dim folderId As FolderId = DirectCast(createdEvent.Id, FolderId)

                                Console.WriteLine("Created folder = " & folderId.ToString())

                            End If
                        ElseIf TypeOf notification.Events(i) Is ModifiedEvent Then

                            Dim modifiedEvent As ModifiedEvent = DirectCast(notification.Events(i), ModifiedEvent)

                            If TypeOf modifiedEvent.Id Is ItemId Then

                                Dim itemId As ItemId = DirectCast(modifiedEvent.Id, ItemId)

                                Console.WriteLine("Modified item = " & itemId.ToString())

                            ElseIf TypeOf modifiedEvent.Id Is FolderId Then

                                Dim folderId As FolderId = DirectCast(modifiedEvent.Id, FolderId)

                                Console.WriteLine("Modified folder = " & folderId.ToString())

                            End If
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
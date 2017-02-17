Imports System
Imports System.IO
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

            Try
                Dim subscription As StreamingSubscription = New StreamingSubscription(StandardFolder.Inbox, EventType.NewMail)

                Dim response As SubscribeResponse = service.Subscribe(subscription)

                Dim responseStream As Stream = service.GetStreamingEvents(response.SubscriptionId, 30)

                While (True)

                    Dim eventsResponse As StreamingEventsResponse = New StreamingEventsResponse(responseStream)

                    Console.WriteLine("New mails:" + eventsResponse.Notifications.Count.ToString())

                    ''exit when connection is closed or error occurs.
                    If (eventsResponse.ConnectionStatus = ConnectionStatus.Closed) Then
                        Exit While
                    ElseIf (eventsResponse.ResponseClass = ResponseClass.Error) Then
                        Console.WriteLine(eventsResponse.XmlMessage)
                        Console.WriteLine(eventsResponse.Message)
                        Exit While
                    End If
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
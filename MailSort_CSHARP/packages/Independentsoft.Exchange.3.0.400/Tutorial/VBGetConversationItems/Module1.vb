Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2013

            Try

                Dim restriction As IsGreaterThanOrEqualTo = New IsGreaterThanOrEqualTo(MessagePropertyPath.ReceivedTime, DateTime.Now.AddMinutes(-15))

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, New ItemShape(ShapeType.Id), restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    Dim request1 As ConversationRequest = New ConversationRequest(response.Items(i).ItemId)

                    Dim conversations As IList(Of ConversationItemResponse) = service.GetConversationItems(request1)

                    For Each conversation As ConversationItemResponse In conversations
                        For Each node As ConversationNode In conversation.Nodes
                            For Each item As Item In node.Items

                                If TypeOf item Is Message Then

                                    Dim message As Message = DirectCast(item, Message)

                                    Console.WriteLine("Subject = " + message.Subject)
                                    Console.WriteLine("ReceivedTime = " + message.ReceivedTime)
                                    Console.WriteLine("----------------------------------------------------------------")

                                End If

                            Next
                        Next
                    Next

                Next

                Console.Read()

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
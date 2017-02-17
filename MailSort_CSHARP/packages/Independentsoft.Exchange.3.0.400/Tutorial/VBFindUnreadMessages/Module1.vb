Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(MessagePropertyPath.IsRead, False)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Message Then

                        Dim message As Message = DirectCast(response.Items(i), Message)

                        Console.WriteLine("Subject = " + message.Subject)
                        Console.WriteLine("ReceivedTime = " + message.ReceivedTime)
                        Console.WriteLine("Body Preview = " + message.BodyPlainText)
                        Console.WriteLine("----------------------------------------------------------------")

                        'If you want to get complete message with entire Body uncomment following line 
                        'message = service.GetMessage(response.Items(i).ItemId)

                        'If you want to set message as read uncomment following lines 
                        'Dim readProperty As New [Property](MessagePropertyPath.IsRead, True)
                        'Dim itemId As ItemId = service.UpdateItem(response.Items(i).ItemId, readProperty)
                    End If
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
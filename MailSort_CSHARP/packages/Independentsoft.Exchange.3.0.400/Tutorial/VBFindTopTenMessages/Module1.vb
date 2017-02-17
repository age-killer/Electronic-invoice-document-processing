Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim order As New List(Of PropertyOrder)()
                order.Add(New PropertyOrder(MessagePropertyPath.ReceivedTime, SortDirection.Descending))

                Dim view As New IndexedPageView()
                view.MaxEntriesReturned = 10

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, order, view)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Message Then

                        Dim message As Message = DirectCast(response.Items(i), Message)

                        Console.WriteLine("Subject = " & message.Subject)
                        Console.WriteLine("ReceivedTime = " & message.ReceivedTime)

                        If message.From IsNot Nothing Then
                            Console.WriteLine("From = " & message.From.Name)
                        End If

                        Console.WriteLine("Body Preview = " & message.BodyPlainText)
                        Console.WriteLine("----------------------------------------------------------------")
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
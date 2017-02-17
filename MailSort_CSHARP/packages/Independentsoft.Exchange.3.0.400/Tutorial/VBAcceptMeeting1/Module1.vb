Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim itemShape As New ItemShape(ShapeType.Id)
                Dim restriction As New IsEqualTo(ItemPropertyPath.ItemClass, ItemClass.MeetingRequest)

                Dim findItemResponse As FindItemResponse = service.FindItem(StandardFolder.Inbox, itemShape, restriction)

                For Each item As Item In findItemResponse.Items

                    Dim acceptItem As New AcceptItem(item.ItemId)

                    Dim response As ItemInfoResponse = service.AcceptMeetingRequest(acceptItem)

                Next

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
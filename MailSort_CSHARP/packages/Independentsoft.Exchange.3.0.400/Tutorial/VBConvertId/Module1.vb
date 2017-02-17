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
                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, New ItemShape(ShapeType.Id))

                For i As Integer = 0 To response.Items.Count - 1

                    Dim sourceId As AlternateId = New AlternateId(IdFormat.EwsId, response.Items(i).ItemId.Id, "John@domain.com")

                    Dim convertIdResponse As ConvertIdResponse = service.ConvertId(sourceId, IdFormat.OwaId)

                    Dim owaId As String = ConvertIdResponse.AlternateId.Id

                    Console.WriteLine("Outlook Web App ItemId:" & owaId)
                Next

                Console.Read()

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " & ex.Message)
                Console.WriteLine("Error: " & ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " & ex.Message)
                Console.Read()
            End Try

        End Sub
    End Class
End Namespace
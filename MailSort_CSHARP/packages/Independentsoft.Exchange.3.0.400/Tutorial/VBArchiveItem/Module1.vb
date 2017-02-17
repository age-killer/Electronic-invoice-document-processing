Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2013

            Try

                Dim items As IList(Of ItemId) = New List(Of ItemId)()

                Dim restriction As IsLessThanOrEqualTo = New IsLessThanOrEqualTo(MessagePropertyPath.ReceivedTime, DateTime.Now.AddMonths(-1))

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, New ItemShape(ShapeType.Id), restriction)

                For Each item As Item In response.Items

                    If TypeOf item Is Message Then

                        items.Add(item.ItemId)

                    End If
                Next

                If (items.Count > 0) Then
                    Dim archiveResponse As IList(Of ItemInfoResponse) = service.ArchiveItem(items, StandardFolder.Inbox)
                End If

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
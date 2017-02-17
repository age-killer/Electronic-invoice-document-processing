Imports System
Imports System.IO
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim itemShape As New ItemShape(ShapeType.Id)
                Dim inboxItems As FindItemResponse = service.FindItem(StandardFolder.Inbox, itemShape)

                For i As Integer = 0 To inboxItems.Items.Count - 1

                    Dim msgFile As Independentsoft.Msg.Message = service.GetMessageFile(inboxItems.Items(i).ItemId)
                    msgFile.Save("c:\test\message" & i & ".msg", True)

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
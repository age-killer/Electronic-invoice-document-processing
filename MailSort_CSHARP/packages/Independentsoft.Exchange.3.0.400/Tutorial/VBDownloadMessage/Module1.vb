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

                    Dim currentItemId As ItemId = inboxItems.Items(i).ItemId

                    Dim message As Message = service.GetMessage(currentItemId)

                    Dim mimeContent As String = message.MimeContent.Text
                    Dim buffer As Byte() = System.Text.Encoding.UTF8.GetBytes(mimeContent)

                    Dim fileName As String = "c:\test\message" & i & ".eml"

                    Dim file As New FileStream(fileName, FileMode.Create)

                    Using file
                        file.Write(buffer, 0, buffer.Length)
                    End Using
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
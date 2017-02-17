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

                Dim fileInfo As New FileInfo("c:\test\message.eml")
                Dim buffer As Byte() = New Byte(fileInfo.Length - 1) {}

                Dim file As New FileStream("c:\test\message.eml", FileMode.Open)

                Using file
                    file.Read(buffer, 0, buffer.Length)
                End Using

                Dim mimeText As String = System.Text.Encoding.UTF8.GetString(buffer)
                Dim mimeContent As New MimeContent(mimeText)

                Dim message As New Message(mimeContent)

                Dim itemId As ItemId = service.CreateItem(message, StandardFolder.Drafts)

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
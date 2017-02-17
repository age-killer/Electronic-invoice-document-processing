Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim message As New Message()
                message.Subject = "Test"
                message.Body = New Body("Body text")
                message.ToRecipients.Add(New Mailbox("John@mydomain.com"))
                message.CcRecipients.Add(New Mailbox("Mark@mydomain.com"))

                Dim itemId As ItemId = service.CreateItem(message)

                Dim response As ItemInfoResponse = service.CopyItem(itemId, StandardFolder.Inbox)

                Dim newItemId As ItemId = response.Items(0).ItemId

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
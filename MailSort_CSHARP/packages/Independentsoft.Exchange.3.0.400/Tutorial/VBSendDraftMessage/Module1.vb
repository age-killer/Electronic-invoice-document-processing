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

                'create message in the Drafts folder 
                Dim itemId1 As ItemId = service.CreateItem(message)

                Dim john As New Mailbox("John@mydomain.com", "John Smith")
                Dim toProperty As New [Property](MessagePropertyPath.ToRecipients, john)

                'Update message. 
                Dim itemId2 As ItemId = service.UpdateItem(itemId1, toProperty)

                'Send message 
                Dim response As ItemInfoResponse = service.Send(itemId2)

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
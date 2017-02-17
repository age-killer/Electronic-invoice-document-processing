Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                'Create message in the Drafts folder 
                Dim message As New Message()
                message.Subject = "Test"
                message.Body = New Body("Body text")
                message.ToRecipients.Add(New Mailbox("John@mydomain.com"))

                Dim messageId As ItemId = service.CreateItem(message)

                'Attach file 
                Dim fileAttachment As New FileAttachment("c:\test\test.docx")
                Dim attachmentId As AttachmentId = service.CreateAttachment(fileAttachment, messageId)

                'Update messageId 
                messageId.ChangeKey = attachmentId.RootItemChangeKey

                'Send message 
                Dim response As ItemInfoResponse = service.Send(messageId)

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
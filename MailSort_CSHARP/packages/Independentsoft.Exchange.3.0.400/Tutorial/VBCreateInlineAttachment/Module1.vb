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
                message.Subject = "Inline"
                message.HideAttachments = True
                message.Body = New Body("<html><body>Here is a message with an inline attachment:<br><img widht=""640"" height=""480"" id=""Picture1"" src=""cid:picture.jpg@123456""></img></body></html>", BodyType.Html)

                Dim itemId As ItemId = service.CreateItem(message)

                Dim fileAttachment As New FileAttachment("c:\test\picture.jpg")
                fileAttachment.ContentId = "picture.jpg@123456"

                Dim attachmentId As AttachmentId = service.CreateAttachment(fileAttachment, itemId)

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
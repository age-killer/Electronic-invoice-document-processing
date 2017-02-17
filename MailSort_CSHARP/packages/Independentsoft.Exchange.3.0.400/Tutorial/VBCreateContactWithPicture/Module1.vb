Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

            Try

                Dim contact As New Contact()
                contact.GivenName = "John"
                contact.Surname = "Smith"
                contact.FileAsMapping = FileAsMapping.LastSpaceFirst
                contact.CompanyName = "Independentsoft"
                contact.BusinessPhone = "555-666-777"
                contact.Email1Address = "john@independentsoft.de"
                contact.Email1DisplayName = "John Smith"
                contact.Email1DisplayAs = "John Smith"
                contact.Email1Type = "SMTP"

                Dim itemId As ItemId = service.CreateItem(contact)

                Dim attachment As New FileAttachment("c:\test\John.jpg")
                attachment.IsContactPhoto = True

                'Add attachment
                Dim attachmentId As AttachmentId = service.CreateAttachment(attachment, itemId)

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
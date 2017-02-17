Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim myPropertyName As New PropertyName("MyCustomProperty", StandardPropertySet.PublicStrings, MapiPropertyType.[String])
                Dim myProperty As New ExtendedProperty(myPropertyName, "value1")

                Dim message As New Message()
                message.Subject = "Test"
                message.Body = New Body("Body text")
                message.ToRecipients.Add(New Mailbox("John@mydomain.com"))
                message.CcRecipients.Add(New Mailbox("Mark@mydomain.com"))
                message.ExtendedProperties.Add(myProperty)

                Dim itemId As ItemId = service.CreateItem(message)

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
Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(ContactPropertyPath.Email1Address, "john@mydomain.com")

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Contacts, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Contact Then

                        Dim itemId As ItemId = response.Items(i).ItemId

                        Dim businessPhoneProperty As New [Property](ContactPropertyPath.BusinessPhone, "555-666-777")

                        itemId = service.UpdateItem(itemId, businessPhoneProperty)

                    End If
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
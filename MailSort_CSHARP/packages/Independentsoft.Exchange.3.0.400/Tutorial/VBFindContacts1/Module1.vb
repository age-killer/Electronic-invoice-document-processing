Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(ContactPropertyPath.GivenName, "Michael")

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Contacts, ContactPropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Contact Then

                        Dim contact As Contact = DirectCast(response.Items(i), Contact)

                        Console.WriteLine("GivenName = " + contact.GivenName)
                        Console.WriteLine("Surname = " + contact.Surname)
                        Console.WriteLine("CompanyName = " + contact.CompanyName)
                        Console.WriteLine("BusinessAddress = " + contact.BusinessAddress)
                        Console.WriteLine("BusinessPhone = " + contact.BusinessPhone)
                        Console.WriteLine("Email1DisplayName = " + contact.Email1DisplayName)
                        Console.WriteLine("Email1Address = " + contact.Email1Address)
                        Console.WriteLine("----------------------------------------------------------------")
                    End If
                Next

                Console.Read()

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
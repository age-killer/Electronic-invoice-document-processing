Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim myPropertyName As New PropertyName("myfield1", StandardPropertySet.PublicStrings, MapiPropertyType.[String])
                Dim restriction As New Exists(myPropertyName)

                Dim propertyPaths As New List(Of PropertyPath)()
                propertyPaths.Add(ContactPropertyPath.GivenName)
                propertyPaths.Add(ContactPropertyPath.Surname)
                propertyPaths.Add(ContactPropertyPath.Email1DisplayName)
                propertyPaths.Add(ContactPropertyPath.Email1Address)
                propertyPaths.Add(myPropertyName)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Contacts, propertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Contact Then

                        Dim contact As Contact = DirectCast(response.Items(i), Contact)

                        Console.WriteLine("GivenName = " + contact.GivenName)
                        Console.WriteLine("Surname = " + contact.Surname)
                        Console.WriteLine("CompanyName = " + contact.CompanyName)
                        Console.WriteLine("Email1DisplayName = " + contact.Email1DisplayName)
                        Console.WriteLine("Email1Address = " + contact.Email1Address)
                        Console.WriteLine("Myfield1 = " + contact.ExtendedProperties(myPropertyName).Value)
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
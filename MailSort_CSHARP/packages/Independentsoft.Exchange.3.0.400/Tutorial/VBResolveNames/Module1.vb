Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim response As ResolveNamesResponse = service.ResolveNames("John")

                For i As Integer = 0 To response.Resolutions.Count - 1

                    Dim resolution As Resolution = response.Resolutions(i)

                    If resolution IsNot Nothing AndAlso resolution.Mailbox IsNot Nothing Then

                        Dim mailbox As Mailbox = resolution.Mailbox

                        Console.WriteLine("Name = " & mailbox.Name)
                        Console.WriteLine("EmailAddress = " & mailbox.EmailAddress)

                    End If

                    If resolution IsNot Nothing AndAlso resolution.Contact IsNot Nothing Then

                        Dim contact As Contact = resolution.Contact

                        Console.WriteLine("GivenName = " & contact.GivenName)
                        Console.WriteLine("EmailAddress = " & contact.Email1Address)

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
Imports System
Imports System.IO
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2013

            Try

                Dim restriction As IsEqualTo = New IsEqualTo(PersonaPropertyPath.Surname, "Smith")

                Dim response As FindPeopleResponse = service.FindPeople(StandardFolder.Contacts, restriction)

                For Each persona As Persona In response.Personas

                    Console.WriteLine("PersonaId:" & persona.PersonaId.ToString())
                    Console.WriteLine("PersonaType:" & persona.PersonaType)
                    Console.WriteLine("DisplayName:" & persona.DisplayName)

                    If persona.PhoneNumber IsNot Nothing Then
                        Console.WriteLine("PhoneNumber:" + persona.PhoneNumber.Number)
                    End If

                    If persona.EmailAddress IsNot Nothing Then
                        Console.WriteLine("EmailAddress:" + persona.EmailAddress.EmailAddress)
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
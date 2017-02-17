Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim response As OutOfOfficeResponse = service.GetOutOfOffice("John@mydomain.com")

                Dim oof As OutOfOffice = response.OutOfOffice

                Console.WriteLine("State = " & oof.State.ToString())

                If oof.InternalReply IsNot Nothing Then
                    Console.WriteLine("InternalReply = " & oof.InternalReply.Message)
                End If

                If oof.ExternalReply IsNot Nothing Then
                    Console.WriteLine("ExternalReply = " & oof.ExternalReply.Message)
                End If

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
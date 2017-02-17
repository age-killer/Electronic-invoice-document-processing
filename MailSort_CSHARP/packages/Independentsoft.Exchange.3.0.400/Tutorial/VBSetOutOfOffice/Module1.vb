Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim oof As New OutOfOffice()
                oof.State = OutOfOfficeState.Scheduled
                oof.Duration = New TimeDuration(DateTime.Today, DateTime.Today.AddDays(10))
                oof.InternalReply = New ReplyBody("I am out of office until " + DateTime.Today.AddDays(10))
                oof.ExternalReply = New ReplyBody("I am out of office until " + DateTime.Today.AddDays(10))

                Dim response As Response = service.SetOutOfOffice(oof, "John@mydomain.com")

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
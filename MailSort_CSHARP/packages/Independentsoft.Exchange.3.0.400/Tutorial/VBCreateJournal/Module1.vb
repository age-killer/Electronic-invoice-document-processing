Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim journal As New Journal()
                journal.Subject = "Test"
                journal.Body = New Body("Body text")
                journal.Type = "Phone call"
                journal.TypeDescription = "Phone call"
                journal.Companies.Add("Independentsoft")
                journal.StartTime = DateTime.Today.AddHours(15)
                journal.EndTime = DateTime.Today.AddHours(16)

                Dim itemId As ItemId = service.CreateItem(journal, StandardFolder.Journal)

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
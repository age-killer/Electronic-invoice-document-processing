Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(JournalPropertyPath.ItemClass, ItemClass.Journal)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Journal, JournalPropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    Dim journal As Journal = DirectCast(response.Items(i), Journal)

                    Console.WriteLine("Subject = " & journal.Subject)
                    Console.WriteLine("Type = " & journal.Type)
                    Console.WriteLine("TypeDescription = " & journal.TypeDescription)
                    Console.WriteLine("StartTime = " & journal.StartTime)
                    Console.WriteLine("EndTime = " & journal.EndTime)
                    Console.WriteLine("Body Preview = " & journal.BodyPlainText)
                    Console.WriteLine("----------------------------------------------------------------")
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
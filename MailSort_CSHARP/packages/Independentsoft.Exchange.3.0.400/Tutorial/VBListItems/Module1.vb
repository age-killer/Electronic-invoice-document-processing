Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim findItemResponse As FindItemResponse = service.FindItem(StandardFolder.Inbox)
                For i As Integer = 0 To findItemResponse.Items.Count - 1

                    Console.WriteLine(findItemResponse.Items(i).Subject)
                    Console.WriteLine(findItemResponse.Items(i).ItemClass)
                    Console.WriteLine(findItemResponse.Items(i).ItemId)
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
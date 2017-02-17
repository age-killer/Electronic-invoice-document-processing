Imports System
Imports System.Net
Imports Independentsoft.Exchange
Imports Independentsoft.Exchange.UnifiedMessaging

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim unifiedMessagingService As New UnifiedMessagingService("https://myserver/UnifiedMessaging/service.asmx", credential)

            Try

                Dim response As String = unifiedMessagingService.PlayOnPhone("AAAAAGsd2rbQLVtLobUGbrq/9IUHAEX2ikn/L8JJtI5WHI0FAW8AAAFXHhsAACxVpEl+KVVLl957wp//x6UAGAetcDUAAA==", "12345")

                Console.WriteLine(response)
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
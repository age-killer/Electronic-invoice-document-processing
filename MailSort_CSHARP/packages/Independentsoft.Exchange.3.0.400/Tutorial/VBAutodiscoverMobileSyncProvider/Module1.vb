Imports System
Imports System.Net
Imports Independentsoft.Exchange
Imports Independentsoft.Exchange.Autodiscover

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim autodiscoverService As New AutodiscoverService("https://myserver/autodiscover/autodiscover.xml", credential)

            Try
                Dim mobileSyncProvider As MobileSyncProvider = autodiscoverService.AutodiscoverMobileSyncProvider("John@mydomain.com")

                Console.WriteLine("Culture = " + mobileSyncProvider.Culture)

                If mobileSyncProvider.User IsNot Nothing Then

                    Dim user As MobileSyncUser = mobileSyncProvider.User

                    Console.WriteLine("DisplayName = " & user.DisplayName)
                    Console.WriteLine("EmailAddress = " & user.EmailAddress)
                End If

                If mobileSyncProvider.Server IsNot Nothing Then

                    Dim server As MobileSyncServer = mobileSyncProvider.Server

                    Console.WriteLine("Server Name = " & server.Name)
                    Console.WriteLine("Server Type = " & server.Type)
                    Console.WriteLine("Server Url  = " & server.Url)
                End If

                Console.Read()

            Catch ex As AutodiscoverException
                Console.WriteLine("Error: " + ex.Message)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " + ex.Message)
                Console.Read()
            End Try

        End Sub
    End Class
End Namespace
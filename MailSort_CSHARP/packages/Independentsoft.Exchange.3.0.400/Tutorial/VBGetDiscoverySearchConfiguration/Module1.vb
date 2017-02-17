Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2013

            Try

                Dim response As GetDiscoverySearchConfigurationResponse = service.GetDiscoverySearchConfiguration("MyDiscoverySearchTest")

                For Each configuration As DiscoverySearchConfiguration In response.DiscoverySearchConfigurations

                    For Each mailbox As SearchableMailbox In configuration.SearchableMailboxes

                        Console.WriteLine("DisplayName:" & mailbox.DisplayName)
                        Console.WriteLine("PrimarySmtpAddress:" & mailbox.PrimarySmtpAddress)
                        Console.WriteLine("ReferenceId:" & mailbox.ReferenceId)
                        Console.WriteLine("-----------------------------------------------")

                    Next

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
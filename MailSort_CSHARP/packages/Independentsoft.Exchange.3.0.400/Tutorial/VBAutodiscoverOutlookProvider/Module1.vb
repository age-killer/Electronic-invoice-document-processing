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
                Dim outlookProvider As OutlookProvider = autodiscoverService.AutodiscoverOutlookProvider("John@mydomain.com")

                If outlookProvider.User IsNot Nothing Then

                    Dim user As User = outlookProvider.User

                    Console.WriteLine("DisplayName = " + user.DisplayName)
                    Console.WriteLine("DeploymentId = " + user.DeploymentId)
                    Console.WriteLine("LegacyDN = " + user.LegacyDN)
                End If

                If outlookProvider.Account IsNot Nothing Then

                    Dim account As Account = outlookProvider.Account

                    Dim exchangeProtocol As ExchangeProtocol = account.ExchangeProtocol
                    Dim webProtocol As WebProtocol = account.WebProtocol

                    If exchangeProtocol IsNot Nothing Then
                        Console.WriteLine("Server = " & exchangeProtocol.Server)
                        Console.WriteLine("ServerDN = " & exchangeProtocol.ServerDN)
                        Console.WriteLine("ServerVersion = " & exchangeProtocol.ServerVersion)
                        Console.WriteLine("ActiveDirectory = " & exchangeProtocol.ActiveDirectory)
                        Console.WriteLine("AuthenticationPackage = " & exchangeProtocol.AuthenticationPackage)
                        Console.WriteLine("MailboxDatabaseLegacyDN = " & exchangeProtocol.MailboxDatabaseLegacyDN)
                        Console.WriteLine("AvailabilityServiceUrl = " & exchangeProtocol.AvailabilityServiceUrl)
                        Console.WriteLine("ExchangeWebServiceUrl = " & exchangeProtocol.ExchangeWebServiceUrl)
                        Console.WriteLine("OfflineAddressBookUrl = " & exchangeProtocol.OfflineAddressBookUrl)
                        Console.WriteLine("OutOfOfficeUrl = " & exchangeProtocol.OutOfOfficeUrl)
                        Console.WriteLine("UnifiedMessagingServiceUrl = " & exchangeProtocol.UnifiedMessagingServiceUrl)
                    End If

                    If webProtocol IsNot Nothing Then
                        If webProtocol.InternalAccess IsNot Nothing Then

                            Dim owaUrls As IList(Of OutlookWebAccessUrl) = webProtocol.InternalAccess.OutlookWebAccessUrls

                            For i As Integer = 0 To owaUrls.Count - 1

                                Console.WriteLine("OWA Url = " & owaUrls(i).Url)

                                For j As Integer = 0 To owaUrls(i).AuthenticationMethods.Count - 1
                                    Console.WriteLine("OWA Authentication = " & owaUrls(i).AuthenticationMethods(j).ToString())
                                Next
                            Next
                        End If

                        If webProtocol.ExternalAccess IsNot Nothing Then

                            Dim owaUrls As IList(Of OutlookWebAccessUrl) = webProtocol.ExternalAccess.OutlookWebAccessUrls

                            For i As Integer = 0 To owaUrls.Count - 1

                                Console.WriteLine("OWA Url = " & owaUrls(i).Url)

                                For j As Integer = 0 To owaUrls(i).AuthenticationMethods.Count - 1
                                    Console.WriteLine("OWA Authentication = " & owaUrls(i).AuthenticationMethods(j).ToString())
                                Next
                            Next
                        End If
                    End If
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
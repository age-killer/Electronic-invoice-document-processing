Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim mailbox As New Mailbox("John@mydomain.com")
                mailbox.RoutingType = "SMTP"

                Dim response1 As ServiceConfigurationResponse = service.GetServiceConfiguration(mailbox, ServiceConfigurationType.MailTips)

                Console.WriteLine("IsMailTipsEnabled=" & response1.MailTipsConfiguration.IsMailTipsEnabled)
                Console.WriteLine("LargeAudienceThreshold=" & response1.MailTipsConfiguration.LargeAudienceThreshold)
                Console.WriteLine("MaxMessageSize=" & response1.MailTipsConfiguration.MaxMessageSize)
                Console.WriteLine("MaxRecipients=" & response1.MailTipsConfiguration.MaxRecipients)
                Console.WriteLine("ShowExternalRecipientCount=" & response1.MailTipsConfiguration.ShowExternalRecipientCount)

                Dim response2 As ServiceConfigurationResponse = service.GetServiceConfiguration(ServiceConfigurationType.UnifiedMessagingConfiguration)

                Console.WriteLine("IsUnifiedMessagingEnabled=" & response2.UnifiedMessagingConfiguration.IsUnifiedMessagingEnabled)
                Console.WriteLine("IsPlayOnPhoneEnabled=" & response2.UnifiedMessagingConfiguration.IsPlayOnPhoneEnabled)
                Console.WriteLine("PlayOnPhoneDialString=" & response2.UnifiedMessagingConfiguration.PlayOnPhoneDialString)

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
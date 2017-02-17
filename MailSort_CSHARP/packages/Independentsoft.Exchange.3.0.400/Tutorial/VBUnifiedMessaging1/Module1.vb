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

                Dim enabled As Boolean = unifiedMessagingService.IsUnifiedMessagingEnabled()
                Dim properties As UnifiedMessagingProperties = unifiedMessagingService.GetUnifiedMessagingProperties()

                Console.WriteLine("Enabled = " & enabled)

                If properties IsNot Nothing Then
                    Console.WriteLine("OutOfOfficeEnabled = " & properties.OutOfOfficeEnabled)
                    Console.WriteLine("MissedCallNotificationEnabled = " & properties.MissedCallNotificationEnabled)
                    Console.WriteLine("PlayOnPhoneDialString = " & properties.PlayOnPhoneDialString)
                    Console.WriteLine("TelephoneAccessFolderEmail = " & properties.TelephoneAccessFolderEmail)
                    Console.WriteLine("TelephoneAccessNumbers = " & properties.TelephoneAccessNumbers)
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
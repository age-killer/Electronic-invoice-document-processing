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

                Dim response As FindMessageTrackingReportResponse = service.FindMessageTrackingReport(MessageTrackingScope.Organization, "mydomain.com", mailbox)

                For Each result As MessageTrackingSearchResult In response.MessageTrackingSearchResults
                    Console.WriteLine("Subject=" & result.Subject)

                    If result.Sender IsNot Nothing Then
                        Console.WriteLine("Sender=" & result.Sender.EmailAddress)
                    End If

                    If result.Sender IsNot Nothing Then
                        Console.WriteLine("PurportedSender=" & result.PurportedSender.EmailAddress)
                    End If

                    Console.WriteLine("PreviousHopServer=" & result.PreviousHopServer)
                    Console.WriteLine("MessageTrackingReportId=" & result.MessageTrackingReportId)
                    Console.WriteLine("SubmittedTime=" & result.SubmittedTime)

                    For Each recipient As Mailbox In result.Recipients
                        Console.WriteLine("Recipient=" & recipient.EmailAddress)
                    Next

                    Console.WriteLine("-----------------------------------------------")
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
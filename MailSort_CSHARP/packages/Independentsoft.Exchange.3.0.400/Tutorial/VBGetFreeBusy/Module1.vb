Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim mailbox As New MailboxData("John@mydomain.com")
                Dim timeZone As New SerializableTimeZone(StandardTimeZone.Berlin)
                Dim freeBusyOptions As New FreeBusyViewOptions(DateTime.Today, DateTime.Today.AddDays(10), FreeBusyViewType.DetailedMerged)

                Dim response As AvailabilityResponse = service.GetAvailability(mailbox, timeZone, freeBusyOptions)

                For i As Integer = 0 To response.FreeBusyResponses.Count - 1

                    Dim freeBusyView As FreeBusyView = response.FreeBusyResponses(i).FreeBusyView

                    For j As Integer = 0 To freeBusyView.CalendarEvents.Count - 1

                        Console.WriteLine("Subject = " + freeBusyView.CalendarEvents(j).Subject)
                        Console.WriteLine("Location = " + freeBusyView.CalendarEvents(j).Location)
                        Console.WriteLine("StartTime = " + freeBusyView.CalendarEvents(j).StartTime)
                        Console.WriteLine("EndTime = " + freeBusyView.CalendarEvents(j).EndTime)
                        Console.WriteLine("-----------------------------------------------------")
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
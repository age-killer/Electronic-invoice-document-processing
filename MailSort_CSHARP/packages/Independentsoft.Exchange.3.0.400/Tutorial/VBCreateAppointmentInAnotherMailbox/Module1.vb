Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim johnMailbox As New Mailbox("John@mydomain.com")
                Dim johnCalendarFolder As New StandardFolderId(StandardFolder.Calendar, johnMailbox)

                Dim appointment As New Appointment()
                appointment.Subject = "Test"
                appointment.Body = New Body("Body text")
                appointment.StartTime = DateTime.Today.AddHours(10)
                appointment.EndTime = DateTime.Today.AddHours(12)
                appointment.Location = "My Office"
                appointment.ReminderIsSet = True
                appointment.ReminderMinutesBeforeStart = 30

                Dim itemId As ItemId = service.CreateItem(appointment, johnCalendarFolder)

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
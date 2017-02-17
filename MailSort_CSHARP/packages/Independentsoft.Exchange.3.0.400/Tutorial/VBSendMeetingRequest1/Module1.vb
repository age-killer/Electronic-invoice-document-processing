Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim appointment As New Appointment()
                appointment.Subject = "Meeting"
                appointment.Body = New Body("Body text")
                appointment.StartTime = DateTime.Today.AddHours(15)
                appointment.EndTime = DateTime.Today.AddHours(16)
                appointment.Location = "Room 123"
                appointment.ReminderIsSet = True
                appointment.ReminderMinutesBeforeStart = 30
                appointment.RequiredAttendees.Add(New Attendee("John@mydomain.com"))
                appointment.OptionalAttendees.Add(New Attendee("Administrator@mydomain.com"))

                Dim itemId As ItemId = service.SendMeetingRequest(appointment)

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
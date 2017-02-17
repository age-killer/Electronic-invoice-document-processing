Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

            Try

                Dim timeZone As TimeZoneDefinition = GetTimeZone(service, "Berlin") ''find time zone based on name
                ''Dim timeZone As TimeZoneDefinition = GetTimeZone(service, "UTC-08:00") ''find time zone based on offset

                'Weekly recurrence. Monday and Tuesday every second week next 3 months. 
                Dim days As New List(Of Independentsoft.Exchange.DayOfWeek)()
                days.Add(Independentsoft.Exchange.DayOfWeek.Monday)
                days.Add(Independentsoft.Exchange.DayOfWeek.Tuesday)

                Dim recurrence As New Recurrence()
                recurrence.Pattern = New WeeklyRecurrencePattern(2, days)
                recurrence.Range = New EndDateRecurrenceRange(DateTime.Today, DateTime.Today.AddMonths(3))

                Dim appointment As New Appointment()
                appointment.Subject = "Every second week, monday and tuesday"
                appointment.Body = New Body("Body text")
                appointment.StartTime = DateTime.Today.AddHours(15)
                appointment.EndTime = DateTime.Today.AddHours(16)
                appointment.Recurrence = recurrence
                appointment.StartTimeZone = timeZone
                appointment.EndTimeZone = timeZone

                Dim itemId As ItemId = service.CreateItem(appointment)

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " + ex.Message)
                Console.WriteLine("Error: " + ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " + ex.Message)
                Console.Read()
            End Try

        End Sub

        Private Shared Function GetTimeZone(ByVal service As Service, ByVal name As String) As TimeZoneDefinition

            Dim Response As GetServerTimeZonesResponse = service.GetServerTimeZones()

            For Each timeZone As TimeZoneDefinition In Response.TimeZoneDefinitions

                If timeZone.Name.IndexOf(name) > 0 Then
                    Return timeZone
                End If
            Next

            Return Nothing

        End Function
    End Class
End Namespace
Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try
                ''Set MeetingTimeZone property. Use only on Exchange 2007
                Dim appointment1 As Appointment = New Appointment()
                appointment1.Subject = "Test1"
                appointment1.Body = New Body("Body text")
                appointment1.StartTime = DateTime.Today.AddHours(10)
                appointment1.EndTime = DateTime.Today.AddHours(11)
                appointment1.MeetingTimeZone = New Independentsoft.Exchange.TimeZone(StandardTimeZone.Berlin)

                Dim itemId1 As ItemId = service.CreateItem(appointment1)

                service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

                Dim timeZone As TimeZoneDefinition = GetTimeZone(service, "Berlin") ''find time zone based on name
                ''Dim timeZone As TimeZoneDefinition = GetTimeZone(service, "UTC-08:00") ''find time zone based on offset

                ''Setting StartTimeZone and EndTimeZone properties. Works on Exchange 2010 and later
                Dim appointment2 As Appointment = New Appointment()
                appointment2.Subject = "Test2"
                appointment2.Body = New Body("Body text")
                appointment2.StartTime = DateTime.Today.AddHours(12)
                appointment2.EndTime = DateTime.Today.AddHours(13)
                appointment2.StartTimeZone = timeZone
                appointment2.EndTimeZone = timeZone

                Dim itemId2 As ItemId = service.CreateItem(appointment2)

                ''Setting TimeZoneContext. It is time zone for entire session
                service.TimeZoneContext = timeZone

                Dim appointment3 As Appointment = New Appointment()
                appointment3.Subject = "Test3"
                appointment3.Body = New Body("Body text")
                appointment3.StartTime = DateTime.Today.AddHours(15)
                appointment3.EndTime = DateTime.Today.AddHours(16)

                Dim itemId3 As ItemId = service.CreateItem(appointment3)

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
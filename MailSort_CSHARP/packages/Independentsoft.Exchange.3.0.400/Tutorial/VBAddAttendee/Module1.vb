Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim appointment As Appointment = service.GetAppointment("AAMkAAAAAAAAYAABClH/p3xPfSLGfvkhFmBB/AAvTWJunAAA=")

                Dim itemId As ItemId = appointment.ItemId

                Dim itemChange As ItemChange = New ItemChange(itemId)
                itemChange.PropertiesToAppend.Add(New [Property](AppointmentPropertyPath.RequiredAttendees, New Attendee("John@mydomain.com")))

                itemId = service.UpdateItem(itemChange, SendMeetingOption.SendToChanged)

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
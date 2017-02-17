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

                ''remove
                Dim itemChange0 As ItemChange = New ItemChange(itemId)
                itemChange0.PropertiesToDelete.Add(AppointmentPropertyPath.RequiredAttendees)
                itemChange0.PropertiesToDelete.Add(AppointmentPropertyPath.OptionalAttendees)

                itemId = service.UpdateItem(itemChange0)

                ''add first
                Dim itemChange1 As ItemChange = New ItemChange(itemId)
                itemChange1.PropertiesToSet.Add(New [Property](AppointmentPropertyPath.RequiredAttendees, New Attendee("John@mydomain.com")))

                itemId = service.UpdateItem(itemChange1)

                ''append others 
                Dim itemChange2 As ItemChange = New ItemChange(itemId)
                itemChange2.PropertiesToAppend.Add(New [Property](AppointmentPropertyPath.RequiredAttendees, New Attendee("Peter@mydomain.com")))

                ''send meeting update.
                itemId = service.UpdateItem(itemChange2, SendMeetingOption.SendToAll)

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
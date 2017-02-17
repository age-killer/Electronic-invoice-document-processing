Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New Contains(MessagePropertyPath.Subject, "test", ContainmentMode.ExactPhrase, ContainmentComparison.IgnoreCase)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Calendar, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Appointment Then

                        Dim requiredAttendees As New [Property](AppointmentPropertyPath.RequiredAttendees, New Attendee("John@mydomain.com"))

                        Dim itemChange As New ItemChange()
                        itemChange.ItemId = response.Items(i).ItemId
                        itemChange.PropertiesToAppend.Add(requiredAttendees)

                        Dim newItemId As ItemId = service.UpdateItem(itemChange, SendMeetingOption.SendToAll)
                    End If

                Next

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
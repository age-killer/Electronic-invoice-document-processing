Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(AppointmentPropertyPath.Subject, "Meeting")

                Dim findItemResponse As FindItemResponse = service.FindItem(StandardFolder.Calendar, restriction)

                Dim startTimeProperty As New [Property](AppointmentPropertyPath.StartTime, DateTime.Today.AddHours(17))
                Dim endTimeProperty As New [Property](AppointmentPropertyPath.EndTime, DateTime.Today.AddHours(20))

                Dim properties As New List(Of [Property])()
                properties.Add(startTimeProperty)
                properties.Add(endTimeProperty)

                If findItemResponse.Items.Count = 1 Then

                    Dim itemId As ItemId = findItemResponse.Items(0).ItemId

                    itemId = service.UpdateItem(itemId, properties, SendMeetingOption.SendToAllAndSaveCopy)

                End If

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
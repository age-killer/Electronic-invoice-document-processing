Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction1 As New IsEqualTo(AppointmentPropertyPath.StartTime, DateTime.Today.AddHours(15))
                Dim restriction2 As New IsEqualTo(AppointmentPropertyPath.EndTime, DateTime.Today.AddHours(16))
                Dim restriction3 As New [And](restriction1, restriction2)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Calendar, restriction3)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Appointment Then

                        Dim itemId As ItemId = response.Items(i).ItemId

                        Dim startTimeProperty As New [Property](AppointmentPropertyPath.StartTime, DateTime.Today.AddHours(18))
                        Dim endTimeProperty As New [Property](AppointmentPropertyPath.EndTime, DateTime.Today.AddHours(19))

                        Dim properties As New List(Of [Property])()
                        properties.Add(startTimeProperty)
                        properties.Add(endTimeProperty)

                        itemId = service.UpdateItem(itemId, properties)

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
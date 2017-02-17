Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction1 As New IsGreaterThanOrEqualTo(AppointmentPropertyPath.StartTime, DateTime.Today)
                Dim restriction2 As New IsLessThanOrEqualTo(AppointmentPropertyPath.EndTime, DateTime.Today.AddDays(1))
                Dim restriction3 As New [And](restriction1, restriction2)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Calendar, AppointmentPropertyPath.AllPropertyPaths, restriction3)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Appointment Then

                        Dim appointment As Appointment = DirectCast(response.Items(i), Appointment)

                        Console.WriteLine("Subject = " + appointment.Subject)
                        Console.WriteLine("StartTime = " + appointment.StartTime)
                        Console.WriteLine("EndTime = " + appointment.EndTime)
                        Console.WriteLine("Body Preview = " + appointment.BodyPlainText)
                        Console.WriteLine("----------------------------------------------------------------")
                    End If
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
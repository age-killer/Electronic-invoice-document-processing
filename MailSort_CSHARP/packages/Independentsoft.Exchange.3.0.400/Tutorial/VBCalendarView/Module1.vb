Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Program
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try
                Dim view As New CalendarView(DateTime.Today, DateTime.Today.AddMonths(1))

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Calendar, AppointmentPropertyPath.AllPropertyPaths, view)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Appointment Then

                        Dim appointment As Appointment = DirectCast(response.Items(i), Appointment)

                        Console.WriteLine("Subject = " & appointment.Subject)
                        Console.WriteLine("StartTime = " & appointment.StartTime)
                        Console.WriteLine("EndTime = " & appointment.EndTime)
                        Console.WriteLine("Body Preview = " & appointment.BodyPlainText)
                        Console.WriteLine("----------------------------------------------------------------")

                        If appointment.InstanceType = InstanceType.Occurrence Then

                            Dim masterId As RecurringMasterItemId = New RecurringMasterItemId(appointment.ItemId.Id, appointment.ItemId.ChangeKey)

                            Dim master As Appointment = service.GetAppointment(masterId)

                        End If
                    End If

                Next

                Console.Read()
            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " & ex.Message)
                Console.WriteLine("Error: " & ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " & ex.Message)
                Console.Read()
            End Try
        End Sub
    End Class
End Namespace

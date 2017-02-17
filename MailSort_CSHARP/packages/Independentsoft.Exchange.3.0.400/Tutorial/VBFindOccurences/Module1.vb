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
                Dim restriction As Exists = New Exists(AppointmentPropertyPath.RecurrenceStart)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Calendar, AppointmentPropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Appointment Then

                        Dim master As Appointment = DirectCast(response.Items(i), Appointment)

                        Dim occurrenceId As OccurrenceItemId = New OccurrenceItemId(master.ItemId.Id, master.ItemId.ChangeKey)

                        ''Get all occurences until exception occurs.
                        For j As Integer = 1 To 999
                            occurrenceId.Index = j

                            Try
                                Dim occurrence As Appointment = service.GetAppointment(occurrenceId)

                                Console.WriteLine("Subject = " & occurrence.Subject)
                                Console.WriteLine("StartTime = " & occurrence.StartTime)
                                Console.WriteLine("EndTime = " & occurrence.EndTime)
                                Console.WriteLine("----------------------------------------------------------------")
                            Catch ex As ServiceRequestException
                                ''ignore exception "Occurrence index is out of recurrence range"
                                Exit For
                            End Try
                        Next
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
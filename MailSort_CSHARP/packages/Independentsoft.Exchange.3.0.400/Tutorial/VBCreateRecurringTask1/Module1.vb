Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                'Daily recurrence. Every second day, next 20 days 
                Dim recurrence As New TaskRecurrence()
                recurrence.Pattern = New DailyRecurrencePattern(2)
                recurrence.Range = New NumberedRecurrenceRange(DateTime.Today, 20)

                Dim task As New Task()
                task.Subject = "Every second day"
                task.Body = New Body("Body text")
                task.Owner = "My Name"
                task.StartDate = DateTime.Today.AddHours(10)
                task.DueDate = DateTime.Today.AddHours(20)
                task.Recurrence = recurrence

                Dim itemId As ItemId = service.CreateItem(task)

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
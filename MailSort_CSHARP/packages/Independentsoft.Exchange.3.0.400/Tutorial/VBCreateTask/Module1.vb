Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim task As New Task()
                task.Subject = "Test"
                task.Body = New Body("Body text")
                task.Owner = "My Name"
                task.StartDate = DateTime.Today
                task.DueDate = DateTime.Today.AddDays(3)
                task.ReminderIsSet = True
                task.ReminderDueBy = DateTime.Today.AddDays(2)

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
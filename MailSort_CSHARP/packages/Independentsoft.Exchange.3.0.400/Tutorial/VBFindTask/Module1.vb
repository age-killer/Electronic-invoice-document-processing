Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(TaskPropertyPath.IsComplete, True)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Tasks, TaskPropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Task Then

                        Dim task As Task = DirectCast(response.Items(i), Task)

                        Console.WriteLine("Subject = " + task.Subject)
                        Console.WriteLine("StartDate = " + task.StartDate)
                        Console.WriteLine("DueDate = " + task.DueDate)
                        Console.WriteLine("Owner = " + task.Owner)
                        Console.WriteLine("Body Preview = " + task.BodyPlainText)
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
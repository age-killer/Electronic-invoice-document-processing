Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(TaskPropertyPath.ItemClass, ItemClass.TaskRequest)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, TaskPropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Task Then

                        Dim taskRequestId As ItemId = response.Items(i).ItemId

                        Dim item As Item = service.GetItem(taskRequestId)

                        Dim itemAttachment As ItemAttachment = DirectCast(service.GetAttachment(item.Attachments(0).Id), ItemAttachment)

                        Dim task As Task = DirectCast(itemAttachment.Item, Task)

                        Console.WriteLine("Subject = " + task.Subject)
                        Console.WriteLine("StartDate = " + task.CommonStartDate)
                        Console.WriteLine("DueDate = " + task.DueDate)
                        Console.WriteLine("Owner = " + task.Owner)
                        Console.WriteLine("Body Preview = " + task.BodyPlainText)
                        Console.WriteLine("----------------------------------------------------------------")

                        'Accept task and create new task in the Tasks folder 
                        Dim taskId As ItemId = service.CreateItem(task, StandardFolder.Tasks)

                        'Delete TaskRequest from the Inbox 
                        service.DeleteItem(taskRequestId)
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
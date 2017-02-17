Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction1 As New IsEqualTo(MessagePropertyPath.LastVerbExecuted, "102") 'LastVerbExecuted.ReplyToSender 
                Dim restriction2 As New IsEqualTo(MessagePropertyPath.LastVerbExecuted, "103") 'LastVerbExecuted.ReplyToAll 
                Dim restriction3 As New IsEqualTo(MessagePropertyPath.LastVerbExecuted, "104") 'LastVerbExecuted.Forward 

                Dim restriction4 As New [Or](restriction1, restriction2, restriction3)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, restriction4)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Message Then

                        Dim message As Message = DirectCast(response.Items(i), Message)

                        Console.WriteLine("Subject = " + message.Subject)
                        Console.WriteLine("LastModifiedTime = " + message.LastModifiedTime)
                        Console.WriteLine("LastVerbExecuted = " + message.LastVerbExecuted.ToString())
                        Console.WriteLine("Body Preview = " + message.BodyPlainText)
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
Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction1 As New IsEqualTo(MessagePropertyPath.FlagStatus, "1") 'FlagStatus.Complete 
                Dim restriction2 As New IsEqualTo(MessagePropertyPath.FlagStatus, "2") 'FlagStatus.Marked 

                Dim restriction3 As New [Or](restriction1, restriction2)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Inbox, MessagePropertyPath.AllPropertyPaths, restriction3)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Message Then

                        Dim message As Message = DirectCast(response.Items(i), Message)

                        Console.WriteLine("Subject = " + message.Subject)
                        Console.WriteLine("FlagStatus = " + message.FlagStatus.ToString())
                        Console.WriteLine("FlagIcon = " + message.FlagIcon.ToString())
                        Console.WriteLine("FlagCompleteTime = " + message.FlagCompleteTime)
                        Console.WriteLine("FlagRequest = " + message.FlagRequest)
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
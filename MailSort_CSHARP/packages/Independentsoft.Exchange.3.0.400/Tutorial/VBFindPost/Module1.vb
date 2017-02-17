Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsGreaterThanOrEqualTo(PostPropertyPath.PostedTime, DateTime.Today.AddMonths(-1))

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Drafts, PostPropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    If TypeOf response.Items(i) Is Post Then

                        Dim post As Post = DirectCast(response.Items(i), Post)

                        Console.WriteLine("Subject = " + post.Subject)
                        Console.WriteLine("PostedTime = " + post.PostedTime)
                        Console.WriteLine("Body Preview = " + post.BodyPlainText)
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
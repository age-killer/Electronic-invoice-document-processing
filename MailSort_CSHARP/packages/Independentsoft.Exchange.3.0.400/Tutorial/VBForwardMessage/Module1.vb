Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(MessagePropertyPath.Subject, "Test")

                Dim inboxResponse As FindItemResponse = service.FindItem(StandardFolder.Inbox, restriction)

                For i As Integer = 0 To inboxResponse.Items.Count - 1

                    Dim currentMessageId As ItemId = inboxResponse.Items(i).ItemId

                    Dim forwardItem As New ForwardItem(currentMessageId)
                    forwardItem.NewBody = New Body("John, please see message below:")
                    forwardItem.ToRecipients.Add(New Mailbox("John@mydomain.com"))

                    Dim response As ItemInfoResponse = service.Forward(forwardItem)
                Next

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
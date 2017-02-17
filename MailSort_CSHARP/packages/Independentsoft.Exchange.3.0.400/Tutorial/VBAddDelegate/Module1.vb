Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "passsword")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim mailbox As New Mailbox("John@mydomain.com")

                Dim delegateUser As New DelegateUser("Mark@mydomain.com")
                delegateUser.CalendarFolderPermissionLevel = DelegateFolderPermissionLevel.Author

                Dim response As DelegateUserResponse = service.AddDelegate(mailbox, delegateUser)

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
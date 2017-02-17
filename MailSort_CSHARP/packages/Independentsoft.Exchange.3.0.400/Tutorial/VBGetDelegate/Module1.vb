Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim mailbox As New Mailbox("John@mydomain.com")

                Dim response As DelegateResponse = service.GetDelegate(mailbox, True)

                Dim delegateUserResponse As IList(Of DelegateUserResponse) = response.DelegateUserResponses
                For i As Integer = 0 To delegateUserResponse.Count - 1

                    Dim user As DelegateUser = delegateUserResponse(i).DelegateUser

                    Console.WriteLine("User = " & user.UserId.DisplayName)
                    Console.WriteLine("CalendarFolderPermissionLevel = " & user.CalendarFolderPermissionLevel.ToString())
                    Console.WriteLine("ContactsFolderPermissionLevel = " & user.ContactsFolderPermissionLevel.ToString())
                    Console.WriteLine("InboxFolderPermissionLevel = " & user.InboxFolderPermissionLevel.ToString())
                    Console.WriteLine("JournalFolderPermissionLevel = " & user.JournalFolderPermissionLevel.ToString())
                    Console.WriteLine("NotesFolderPermissionLevel = " & user.NotesFolderPermissionLevel.ToString())
                    Console.WriteLine("ReceiveCopiesOfMeetingMessages = " & user.ReceiveCopiesOfMeetingMessages.ToString())
                    Console.WriteLine("TasksFolderPermissionLevel = " & user.TasksFolderPermissionLevel.ToString())
                    Console.WriteLine("ViewPrivateItems = " & user.ViewPrivateItems.ToString())
                    Console.WriteLine("------------------------------------------------------------------------")
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
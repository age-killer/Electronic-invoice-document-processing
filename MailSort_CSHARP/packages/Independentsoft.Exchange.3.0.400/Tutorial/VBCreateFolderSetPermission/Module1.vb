Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try
                Dim permission1 As New CalendarPermission()
                permission1.UserId = New UserId("Mark@mydomain.com")
                permission1.Level = CalendarPermissionLevel.Author

                Dim calFolder As New CalendarFolder("TestCalendar")
                calFolder.PermissionSet = New CalendarPermissionSet()
                calFolder.PermissionSet.Permissions.Add(permission1)

                Dim folder1Id As FolderId = service.CreateFolder(calFolder, StandardFolder.Calendar)

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
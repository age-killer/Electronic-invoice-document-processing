Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim permission1 As New Permission()
                permission1.UserId = New UserId("Mark@mydomain.com")

                permission1.Level = PermissionLevel.[Custom]
                permission1.CanCreateItems = True
                permission1.CanCreateSubFolders = True

                Dim permissionSet As New PermissionSet()
                permissionSet.Permissions.Add(permission1)

                Dim permissionSetProperty As New [Property](FolderPropertyPath.PermissionSet, permissionSet)

                Dim contactsFolder As Folder = service.GetFolder(StandardFolder.Contacts)

                Dim change As New FolderChange(contactsFolder.FolderId, permissionSetProperty)

                service.UpdateFolder(change)

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
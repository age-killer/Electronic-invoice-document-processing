Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim myCustomPropertyName As New PropertyName("myproperty", StandardPropertySet.PublicStrings, MapiPropertyType.[String])

                Dim myCustomProperty As New [Property](myCustomPropertyName, "value1")
                Dim commentProperty As New [Property](FolderPropertyPath.Comment, "Folder description text.")

                'Get only FolderId property of the Inbox folder 
                Dim inboxFolder As Folder = service.GetFolder(StandardFolder.Inbox, ShapeType.Id)

                Dim newFolderId1 As FolderId = service.UpdateFolder(inboxFolder.FolderId, myCustomProperty)
                Dim newFolderId2 As FolderId = service.UpdateFolder(inboxFolder.FolderId, commentProperty)

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
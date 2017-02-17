Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                'Get all folder properties 
                Dim inboxFolder1 As Folder = service.GetFolder(StandardFolder.Inbox)

                'Get only FolderId property of the Inbox folder 
                Dim inboxFolder2 As Folder = service.GetFolder(StandardFolder.Inbox, ShapeType.Id)

                Dim propertyPaths As New List(Of PropertyPath)()
                propertyPaths.Add(FolderPropertyPath.DisplayName)
                propertyPaths.Add(FolderPropertyPath.FolderClass)
                propertyPaths.Add(FolderPropertyPath.Comment)

                'Get specified properties 
                Dim inboxFolder3 As Folder = service.GetFolder(StandardFolder.Inbox, propertyPaths)

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
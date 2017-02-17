Imports System
Imports System.Net
Imports System.Collections
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1

        Private Shared folderTable As New Hashtable()

        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try
                Dim findFolderResponse As FindFolderResponse = service.FindFolder(StandardFolder.MailboxRoot, FolderPropertyPath.AllPropertyPaths, FolderQueryTraversal.Deep)

                For i As Integer = 0 To findFolderResponse.Folders.Count - 1
                    Dim folderId As String = findFolderResponse.Folders(i).FolderId.Id
                    Dim folder As Folder = findFolderResponse.Folders(i)

                    folderTable.Add(folderId, folder)
                Next

                For Each folderId As String In folderTable.Keys
                    Dim folder As Folder = DirectCast(folderTable(folderId), Folder)

                    Dim folderPath As String = GetFolderPath(folder)

                    Console.WriteLine(folder.DisplayName)
                    Console.WriteLine(folder.FolderClass)
                    Console.WriteLine(folder.CreationTime)
                    Console.WriteLine(folderPath)
                    Console.WriteLine("--------------------------------------------------------------------")
                Next

                Console.Read()

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " & ex.Message)
                Console.WriteLine("Error: " & ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " & ex.Message)
                Console.Read()
            End Try
        End Sub

        Private Shared Function GetFolderPath(ByVal folder As Folder) As String
            Dim path As String = ""

            path = (path & "/") + folder.DisplayName

            Dim parentId As String = folder.ParentId.Id

            If folderTable(parentId) IsNot Nothing Then
                Dim parentFolder As Folder = DirectCast(folderTable(parentId), Folder)

                Return GetFolderPath(parentFolder) + path
            End If

            Return path
        End Function

    End Class
End Namespace

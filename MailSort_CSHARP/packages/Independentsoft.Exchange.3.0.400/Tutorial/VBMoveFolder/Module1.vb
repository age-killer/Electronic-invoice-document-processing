Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(FolderPropertyPath.DisplayName, "Test")

                Dim findFolderResponse As FindFolderResponse = service.FindFolder(StandardFolder.MailboxRoot, restriction)

                For i As Integer = 0 To findFolderResponse.Folders.Count - 1

                    Dim currentFolderId As FolderId = findFolderResponse.Folders(i).FolderId

                    Dim newFolderId As FolderId = service.MoveFolder(currentFolderId, StandardFolder.Drafts)
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
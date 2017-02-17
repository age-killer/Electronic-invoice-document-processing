Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try
                Dim folder1 As New Folder("Test1")
                Dim folder2 As New Folder("Test2")

                Dim folder1Id As FolderId = service.CreateFolder(folder1, StandardFolder.Drafts)
                Dim folder2Id As FolderId = service.CreateFolder(folder2, StandardFolder.Drafts)

                Dim response1 As Response = service.DeleteFolder(folder1Id) 'hard delete 
                Dim response2 As Response = service.DeleteFolder(folder2Id, DeleteType.MoveToDeletedItems)

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
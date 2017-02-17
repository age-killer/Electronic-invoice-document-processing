Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(FolderPropertyPath.FolderClass, FolderClass.ContactFolder)

                Dim findFolderResponse As FindFolderResponse = service.FindFolder(StandardFolder.MailboxRoot, restriction, FolderQueryTraversal.Deep)
                For i As Integer = 0 To findFolderResponse.Folders.Count - 1

                    Console.WriteLine(findFolderResponse.Folders(i).DisplayName)
                    Console.WriteLine(findFolderResponse.Folders(i).FolderId)
                    Console.WriteLine("----------------------------------------------------------")
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
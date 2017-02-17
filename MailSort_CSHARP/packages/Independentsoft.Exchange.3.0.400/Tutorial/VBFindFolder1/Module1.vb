Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction1 As New IsGreaterThanOrEqualTo(FolderPropertyPath.CreationTime, DateTime.Today)
                Dim restriction2 As New IsLessThan(FolderPropertyPath.CreationTime, DateTime.Today.AddDays(1))
                Dim restriction3 As New [And](restriction1, restriction2)

                Dim findFolderResponse As FindFolderResponse = service.FindFolder(StandardFolder.MailboxRoot, FolderPropertyPath.AllPropertyPaths, restriction3, FolderQueryTraversal.Deep)

                For i As Integer = 0 To findFolderResponse.Folders.Count - 1

                    Console.WriteLine(findFolderResponse.Folders(i).DisplayName)
                    Console.WriteLine(findFolderResponse.Folders(i).FolderClass)
                    Console.WriteLine(findFolderResponse.Folders(i).CreationTime)
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
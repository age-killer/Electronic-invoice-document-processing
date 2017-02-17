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

                'Create a folder in the mailbox root
                Dim folder1Id As FolderId = service.CreateFolder(folder1, StandardFolder.MailboxRoot)

                'Create a subfolder of the Test1
                Dim folder2Id As FolderId = service.CreateFolder(folder2, folder1Id)

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
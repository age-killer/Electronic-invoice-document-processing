Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "parameter")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                'Initial synscronization. It will recursively retrieve all existing folders from the mailbox 
                Dim response As SyncFoldersResponse = service.SyncFolders(StandardFolder.MailboxRoot)

                'Keep synchronization state in order to use it in next synchronizations 
                Dim state As String = response.State

                Dim allExistingFolders As IList(Of Folder) = response.CreatedFolders

                'Create new folder 
                Dim folderId As FolderId = service.CreateFolder("test", StandardFolder.MailboxRoot)

                'Synchronize folders again but now include the state parameter in order to find changes from last synchronization 
                Dim newResponse As SyncFoldersResponse = service.SyncFolders(StandardFolder.MailboxRoot, state)

                Dim newFolders As IList(Of Folder) = newResponse.CreatedFolders
                Dim deletedFolders As IList(Of FolderId) = newResponse.DeletedFolders
                Dim updatedFolders As IList(Of Folder) = newResponse.UpdatedFolders

                'Display new created folders 
                For i As Integer = 0 To newFolders.Count - 1

                    Console.WriteLine(newFolders(i).DisplayName)
                    Console.WriteLine(newFolders(i).FolderId)
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
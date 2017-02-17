Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            ''works on Exchange 2013 / Office365 and higher
            service.RequestServerVersion = RequestServerVersion.Exchange2013

            Try
                Dim subfolders As IList(Of Folder) = New List(Of Folder)()
                subfolders.Add(New Folder("Folder1")) ''Folder1 is under Inbox folder
                subfolders.Add(New Folder("Folder2")) ''Folder2 is under Folder1
                subfolders.Add(New Folder("Folder3")) ''Folder3 is under Folder2

                Dim response As IList(Of FolderInfoResponse) = service.CreateFolderPath(subfolders, StandardFolder.Inbox)

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
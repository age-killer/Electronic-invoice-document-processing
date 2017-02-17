Imports System
Imports System.IO
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

            Try

                Dim list As CategoryList = service.GetCategoryList()

                Dim myCategory As Category = New Category("Steel Category", CategoryColor.DarkSteel, KeyboardShortcut.CtrlF7)
                myCategory.Guid = "{fac7e643-c161-4f81-b394-898943335ad4}"

                list.Categories.Add(myCategory)

                service.UpdateCategoryList(List)

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
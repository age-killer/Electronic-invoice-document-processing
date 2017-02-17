Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim note As New Note()
                note.Subject = "My test note"
                note.Body = New Body("My test note")
                note.Color = NoteColor.Green
                note.IconColor = NoteColor.Green
                note.Height = 200
                note.Width = 300
                note.Left = 400
                note.Top = 200

                Dim itemId As ItemId = service.CreateItem(note, StandardFolder.Notes)

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
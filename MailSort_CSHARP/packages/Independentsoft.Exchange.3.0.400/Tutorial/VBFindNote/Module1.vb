Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim restriction As New IsEqualTo(NotePropertyPath.ItemClass, ItemClass.Note)

                Dim response As FindItemResponse = service.FindItem(StandardFolder.Notes, NotePropertyPath.AllPropertyPaths, restriction)

                For i As Integer = 0 To response.Items.Count - 1

                    Dim note As Note = DirectCast(response.Items(i), Note)

                    Console.WriteLine("Subject = " & note.Subject)
                    Console.WriteLine("Width = " & note.Width)
                    Console.WriteLine("Height = " & note.Height)
                    Console.WriteLine("Left = " & note.Left)
                    Console.WriteLine("Top = " & note.Top)
                    Console.WriteLine("Color = " & note.Color.ToString())
                    Console.WriteLine("IconColor = " & note.IconColor.ToString())
                    Console.WriteLine("Body Preview = " & note.BodyPlainText)
                    Console.WriteLine("----------------------------------------------------------------")
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
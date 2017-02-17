Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "passsword")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            Try

                Dim mailboxes As New List(Of MailboxData)()
                mailboxes.Add(New MailboxData("John@mydomain.com"))
                mailboxes.Add(New MailboxData("Peter@mydomain.com"))
                mailboxes.Add(New MailboxData("Mark@mydomain.com"))

                Dim timeZone As New SerializableTimeZone(StandardTimeZone.Berlin)
                Dim suggestionsViewOptions As New SuggestionsViewOptions(DateTime.Today.AddDays(1), DateTime.Today.AddDays(2), 60)

                Dim response As AvailabilityResponse = service.GetAvailability(mailboxes, timeZone, suggestionsViewOptions)

                If response.SuggestionsResponse IsNot Nothing Then

                    Dim suggestionDays As IList(Of SuggestionDay) = response.SuggestionsResponse.SuggestionDays

                    For i As Integer = 0 To suggestionDays.Count - 1

                        Console.WriteLine("Suggestion Day = " & suggestionDays(i).Date)

                        For j As Integer = 0 To suggestionDays(i).Suggestions.Count - 1

                            Dim suggestion As Suggestion = suggestionDays(i).Suggestions(j)

                            Console.WriteLine("MeetingTime = " & suggestion.MeetingTime)
                            Console.WriteLine("Quality = " & suggestion.Quality.ToString())
                        Next

                        Console.WriteLine("---------------------------------------------------------")
                    Next
                End If

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
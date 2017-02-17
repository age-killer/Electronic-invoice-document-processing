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

                Dim rule As Rule = New Rule()
                rule.DisplayName = "TestRule"

                rule.Conditions = New RuleConditions()
                rule.Conditions.ContainsSubjectStrings.Add("REMOVEME")

                rule.Actions = New RuleActions()
                rule.Actions.Delete = True

                Dim response As RuleOperationResponse = service.CreateRule(rule)

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
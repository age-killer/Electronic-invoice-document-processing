Imports System
Imports System.Net
Imports Independentsoft.Exchange

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim credential As New NetworkCredential("username", "password")
            Dim service As New Service("https://myserver/ews/Exchange.asmx", credential)

            service.RequestServerVersion = RequestServerVersion.Exchange2010SP1

            Try

                Dim settings As RetentionSettings = service.GetRetentionSettings(StandardFolder.Inbox)

                Dim myRetentionPolicy As RetentionPolicyTag = New RetentionPolicyTag()

                myRetentionPolicy.Id = Guid.NewGuid().ToString()
                myRetentionPolicy.ObjectGuid = Guid.NewGuid().ToString()
                myRetentionPolicy.DisplayName = "Delete after 3 months"
                myRetentionPolicy.RetentionAction = RetentionAction.DeleteAndAllowRecovery
                myRetentionPolicy.Period = 90 ''90 days
                myRetentionPolicy.Type = ElcFolderType.Personal
                myRetentionPolicy.IsArchive = False
                myRetentionPolicy.IsVisible = True
                myRetentionPolicy.OptedInto = True

                myRetentionPolicy.ContentSettings = New ContentSettings()
                myRetentionPolicy.ContentSettings.Id = Guid.NewGuid().ToString()
                myRetentionPolicy.ContentSettings.MessageClass = "*"
                myRetentionPolicy.ContentSettings.Period = 90
                myRetentionPolicy.ContentSettings.RetentionAction = RetentionAction.DeleteAndAllowRecovery

                settings.RetentionPolicyTags.Add(myRetentionPolicy)

                Dim response As Response = service.UpdateRetentionSettings(settings, StandardFolder.Inbox)

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
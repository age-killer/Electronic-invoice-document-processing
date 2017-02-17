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
                Dim response As GetRulesResponse = service.GetRules()

                For Each rule As Rule In response.Rules

                    Console.WriteLine("Id:" & rule.Id)
                    Console.WriteLine("DisplayName:" & rule.DisplayName)

                    If rule.Actions IsNot Nothing Then
                        Console.WriteLine("Actions")
                        Console.WriteLine("AssignCategories count:" & rule.Actions.AssignCategories.Count)
                        Console.WriteLine("Delete:" & rule.Actions.Delete)
                        Console.WriteLine("ForwardAsAttachmentToRecipients count:" & rule.Actions.ForwardAsAttachmentToRecipients.Count)
                        Console.WriteLine("ForwardToRecipients count:" & rule.Actions.ForwardToRecipients.Count)
                        Console.WriteLine("MarkAsRead:" & rule.Actions.MarkAsRead)
                        Console.WriteLine("MarkImportance:" & rule.Actions.MarkImportance)
                        Console.WriteLine("PermanentDelete:" & rule.Actions.PermanentDelete)
                        Console.WriteLine("RedirectToRecipients count:" & rule.Actions.RedirectToRecipients.Count)
                        Console.WriteLine("SendSMSAlertToRecipients count:" & rule.Actions.SendSMSAlertToRecipients.Count)
                        Console.WriteLine("StopProcessingRules:" & rule.Actions.StopProcessingRules)
                        Console.WriteLine("-----------------------------------------------------------------------")
                    End If

                    If rule.Conditions IsNot Nothing Then
                        Console.WriteLine("Conditions")
                        Console.WriteLine("Categories count:" & rule.Conditions.Categories.Count)
                        Console.WriteLine("ContainsBodyStrings count:" & rule.Conditions.ContainsBodyStrings.Count)
                        Console.WriteLine("ContainsHeaderStrings count:" & rule.Conditions.ContainsHeaderStrings.Count)
                        Console.WriteLine("ContainsRecipientStrings count:" & rule.Conditions.ContainsRecipientStrings.Count)
                        Console.WriteLine("ContainsSenderStrings count:" & rule.Conditions.ContainsSenderStrings.Count)
                        Console.WriteLine("ContainsSubjectOrBodyStrings count:" & rule.Conditions.ContainsSubjectOrBodyStrings.Count)
                        Console.WriteLine("ContainsSubjectStrings count:" & rule.Conditions.ContainsSubjectStrings.Count)
                        Console.WriteLine("FlaggedForAction:" & rule.Conditions.FlaggedForAction)
                        Console.WriteLine("FromAddresses count:" & rule.Conditions.FromAddresses.Count)
                        Console.WriteLine("FromConnectedAccounts count:" & rule.Conditions.FromConnectedAccounts.Count)
                        Console.WriteLine("HasAttachments:" & rule.Conditions.HasAttachments)
                        Console.WriteLine("Importance:" & rule.Conditions.Importance)
                        Console.WriteLine("IsApprovalRequest:" & rule.Conditions.IsApprovalRequest)
                        Console.WriteLine("IsAutomaticForward:" & rule.Conditions.IsAutomaticForward)
                        Console.WriteLine("IsAutomaticReply:" & rule.Conditions.IsAutomaticReply)
                        Console.WriteLine("IsEncrypted:" & rule.Conditions.IsEncrypted)
                        Console.WriteLine("IsMeetingRequest:" & rule.Conditions.IsMeetingRequest)
                        Console.WriteLine("IsMeetingResponse:" & rule.Conditions.IsMeetingResponse)
                        Console.WriteLine("IsNDR:" & rule.Conditions.IsNDR)
                        Console.WriteLine("IsPermissionControlled:" & rule.Conditions.IsPermissionControlled)
                        Console.WriteLine("IsReadReceipt:" & rule.Conditions.IsReadReceipt)
                        Console.WriteLine("IsSigned:" & rule.Conditions.IsSigned)
                        Console.WriteLine("IsVoiceMail:" & rule.Conditions.IsVoiceMail)
                        Console.WriteLine("ItemClasses count:" & rule.Conditions.ItemClasses.Count)
                        Console.WriteLine("MessageClassifications count:" & rule.Conditions.MessageClassifications.Count)
                        Console.WriteLine("NotSentToMe:" & rule.Conditions.NotSentToMe)
                        Console.WriteLine("Sensitivity:" & rule.Conditions.Sensitivity)
                        Console.WriteLine("SentCcMe:" & rule.Conditions.SentCcMe)
                        Console.WriteLine("SentOnlyToMe:" & rule.Conditions.SentOnlyToMe)
                        Console.WriteLine("SentToAddresses count:" & rule.Conditions.SentToAddresses.Count)
                        Console.WriteLine("SentToMe:" & rule.Conditions.SentToMe)
                        Console.WriteLine("SentToOrCcMe:" & rule.Conditions.SentToOrCcMe)
                        Console.WriteLine("-----------------------------------------------------------------------")
                    End If
                Next

                Console.Read()

            Catch ex As ServiceRequestException
                Console.WriteLine("Error: " & ex.Message)
                Console.WriteLine("Error: " & ex.XmlMessage)
                Console.Read()
            Catch ex As WebException
                Console.WriteLine("Error: " & ex.Message)
                Console.Read()
            End Try

        End Sub
    End Class
End Namespace
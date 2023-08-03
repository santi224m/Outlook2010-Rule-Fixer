Attribute VB_Name = "UpdateFolder"
Sub UpdateFolder()
    Dim colRules As Outlook.Rules
    Dim oRule As Outlook.Rule
    Dim oMoveAction As Outlook.MoveOrCopyRuleAction
    Dim oInbox As Outlook.Folder
    Dim oMoveTarget As Outlook.Folder
    Dim oRecepient As Outlook.recipient
    Dim strSpecifiedFolder As String
    Dim strDestinationFolder As String
    Dim strAddress As String
    Dim dictAddressToFolder ' Dictionary
    Dim oNewRule As Outlook.Rule
    Dim colRuleActions As Outlook.RuleActions
    Dim oNewMoveRuleAction As Outlook.MoveOrCopyRuleAction
    Dim oNewFromCondition As Outlook.ToOrFromRuleCondition
    Dim strTempName As String
    
    ' Define dictAddressToFolder dictionary (Add Dictionary here)
    Set dictAddressToFolder = CreateObject("Scripting.Dictionary")
    
    ' Set the inbox folder
    Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
    
    ' Get rules from Session.DefaultStore object
    Set colRules = Application.Session.DefaultStore.GetRules()
    
    ' Loop through all rules in colRules collection
    For Each oRule In colRules
        ' Select rules that contain a MoveOrCopyRuleAction
        If oRule.Actions.MoveToFolder.Enabled Then
            ' Filter rules that are missing moveTo folder
            On Error Resume Next
            strSpecifiedFolder = oRule.Actions.MoveToFolder.Folder
            If Err.Number <> 0 Then
                Err.Clear
                ' Create new rule
                Set oNewRule = colRules.Create(oRule.Name & " (fixed)", olRuleReceive)
                Set oNewFromCondition = oNewRule.Conditions.From
                ' Get recepient email address
                For Each oRecepient In oRule.Conditions.From.Recipients
                    strAddress = oRecepient.Address
                    ' Specify oMoveTarget using dictAddressToFolder dict
                    strDestinationFolder = dictAddressToFolder.Item(strAddress)
                    Set oMoveTarget = oInbox.Folders(strDestinationFolder)
                    ' Add recipient to condition
                    With oNewFromCondition
                        .Enabled = True
                        .Recipients.Add (oRecepient.Address)
                        .Recipients.ResolveAll
                    End With
                Next oRecepient
                Set oNewMoveRuleAction = oNewRule.Actions.MoveToFolder
                ' Specify action for new rule
                With oNewMoveRuleAction
                    .Enabled = True
                    .Folder = oMoveTarget
                End With
                oNewRule.Actions.Stop.Enabled = True
                ' Specify the MoveToFolder object
                Set oMoveAction = oRule.Actions.MoveToFolder
                With oMoveAction
                    .Enabled = True
                    .Folder = oMoveTarget
                End With
                ' Set rule to Enabled
                oRule.Enabled = True
                Debug.Print "Modified " & oRule.Name & "..."
                strTempName = oRule.Name
                colRules.Remove (oRule.Name)
                oNewRule.Name = strTempName
            End If
        End If
    Next oRule
    ' Save changes to rules
    colRules.Save
End Sub

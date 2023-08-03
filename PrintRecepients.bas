Attribute VB_Name = "PrintRecepients"
Sub PrintRecepients()
    Dim colRules As Outlook.Rules
    Dim oRule As Outlook.Rule
    Dim colRuleActions As Outlook.RuleActions
    Dim oMoveAction As Outlook.MoveOrCopyRuleAction
    Dim oFromCondition As Outlook.ToOrFromRuleCondition
    Dim oMoveTarget As Outlook.Folder
    Dim colRecepients As Outlook.Recipients
    Dim oRecepient As Outlook.recipient
    Dim sFolder As String
    
    ' Get rules from Session.DefaultStore object
    Set colRules = Application.Session.DefaultStore.GetRules()
    
    ' Loop through all rules in colRules collection
    For Each oRule In colRules
        ' Select rules that contain a MoveOrCopyRuleAction
        If oRule.Actions.MoveToFolder.Enabled Then
            ' Filter rules that are missing moveTo folder
            On Error Resume Next
            sFolder = oRule.Actions.MoveToFolder.Folder
            If Err.Number <> 0 Then
                Err.Clear
                ' Print recpient address
                Set colRecepients = oRule.Conditions.From.Recipients
                For Each oRecepient In colRecepients
                    Debug.Print oRecepient.Address
                Next oRecepient
            End If
        End If
    Next oRule
End Sub

Attribute VB_Name = "PrintFolders"
Sub PrintFolders()
    Dim oInbox As Outlook.Folder
    Dim colFolders As Outlook.Folders
    Dim oChildFolder As Outlook.Folder
    
    ' Set Inbox folder
    Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox)
    ' Create Inbox children collection
    Set colFolders = oInbox.Folders
    
    'Iterater through folders in colFolders collection and print names
    For Each oChildFolder In colFolders
        Debug.Print oChildFolder.Name
    Next oChildFolder
End Sub

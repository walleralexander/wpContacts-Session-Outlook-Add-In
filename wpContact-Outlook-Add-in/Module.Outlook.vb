
Module ModuleOutlook
    Public objApp As Outlook.Application
    Public outlookNameSpace As Outlook.NameSpace
    Public Function CheckOutlookFolderExists(myfolder As String, Optional tag As String = "unknown") As Boolean
        Dim outlookObj As Outlook._Application = New Outlook.Application
        Dim contactFolder As Outlook.MAPIFolder = outlookObj.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts)
        Dim olMandatare As Outlook.MAPIFolder
        Try
            olMandatare = contactFolder.Folders(myfolder)
            MyLog($"CheckOutlookFolderExists/{tag}: Found folder '{myfolder}' Anzahl:{olMandatare.Items.Count}")
            Return True
        Catch ex As Exception
            MyLog($"CheckOutlookFolderExists/{tag}: {ex.Message}", "error")
            Return False
        Finally
            outlookObj = Nothing
            contactFolder = Nothing
            olMandatare = Nothing
        End Try
    End Function
    Public Function CreateOutlookFolder(myfolder) As Boolean
        Dim outlookObj As Outlook._Application = New Outlook.Application
        Dim contactFolder As Outlook.MAPIFolder = outlookObj.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
        Dim olMandatare As Outlook.MAPIFolder
        Try
            MyLog("CreateOutlookFolder: Creating Oulook folder '" & myfolder & "'")
            Dim newfolder As Outlook.Folder = contactFolder.Folders.Add(myfolder, Outlook.OlDefaultFolders.olFolderContacts)
            newfolder.ShowAsOutlookAB = True
            Return True
        Catch ex As Exception
            MyLog($"CreateOutlookFolder: {ex.Message}", "error")
            Return False
        Finally
            outlookObj = Nothing
            contactFolder = Nothing
            olMandatare = Nothing
        End Try
    End Function
    Private Sub DeleteContact(lastName As String, firstName As String)
        Dim contact As Outlook.ContactItem = TryCast(objApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Items.Find(String.Format("[LastName]='{0}' AND [FirstName]='{1}'", lastName, firstName)), Outlook.ContactItem)
        If (contact IsNot Nothing) Then
            contact.Delete()
        End If
    End Sub
    Private Sub FindContactEmailByName(firstName As String, lastName As String)
        Dim contactFolder As Outlook.MAPIFolder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
        Dim contactItems As Outlook.Items = contactFolder.Folders("Kontakte Mandatare").Items

        Try
            Dim contact As Outlook.ContactItem = CType(contactItems.Find(String.Format("[FirstName]='{0}' and [LastName]={1}", firstName, lastName)), Outlook.ContactItem)
            If contact IsNot Nothing Then
                contact.Display()
            Else
                MyLog("FindContactEmailByName: The contact information was not found.")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Module

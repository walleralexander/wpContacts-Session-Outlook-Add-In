'### LIZENZ / LICENSE https://github.com/walleralexander/wpContacts-Session-Outlook-Add-In/blob/master/LICENSE.txt

'    Copyright(C) 2021  Alexander Waller
'
'    This program Is free software: you can redistribute it And/Or modify
'    it under the terms Of the GNU Affero General Public License As
'    published by the Free Software Foundation, either version 3 Of the
'    License, Or (at your option) any later version.
'
'    This program Is distributed In the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty Of
'    MERCHANTABILITY Or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License For more details.
'
'    You should have received a copy Of the GNU Affero General Public License
'    along with this program.  If Not, see < https: //www.gnu.org/licenses/>.
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
            MyError(ex, $"CheckOutlookFolderExists/{tag}: {ex.Message}")
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
            MyError(ex, "CreateOutlookFolder")
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

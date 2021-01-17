Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Data
'Imports Microsoft.Exchange.WebServices.Data

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

Public Class FormSync
    '### ToDo: Eigene Progressbar einführen bei der man die Farbe ändern kann.
    Sub RunSync()
        Try
            Cursor.Current = Cursors.WaitCursor
            lblSyncStatus.Text = "Synchronisation gestartet ..."
            Refresh()
            Dim olFolder As String = My.Settings.OlFolder
            cbxFolder.Text = cbxFolder.Text & " '" & olFolder & "'"
            Dim olContactGroupPrefix As String = My.Settings.OlContactGroupPrefix
            Dim outlookObj As Outlook._Application = New Outlook.Application
            Dim contactFolder As Outlook.MAPIFolder = outlookObj.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
            cbxFolder.Checked = True
            Dim olMandatare As Outlook.MAPIFolder = contactFolder.Folders(olFolder)
            MyLog($"Found folder '{olFolder}' Anzahl:{olMandatare.Items.Count}")
            lblOutlookItems.Text = olMandatare.Items.Count
            Refresh()
            cbxOutlook.Checked = True

            Dim con As SqlConnection
            con = SqlBuildConnection(My.Settings.SQLDataSource, My.Settings.SQLInitialCatalog, My.Settings.SQLUserID, My.Settings.SQLPassword)
            If con Is Nothing Then
                Cursor.Current = Cursors.Default
                Exit Sub
            End If
            cbxSQLConnection.Checked = True
            Dim SessionContacts As DataTable = ReadContactsFromSession(con)
            cbxReadContactsFromSession.Checked = True
            lblReadContactsFromSession.Text = SessionContacts.Rows.Count
            pgbReadContactsFromSession.Maximum = SessionContacts.Rows.Count
            My.Settings.AnzSessionMitarbeiter = SessionContacts.Rows.Count

            Dim SessionGremien As DataTable = ReadGremienFromSession(con)
            cbxReadGremienFromSession.Checked = True
            lblReadGremienFromSession.Text = SessionGremien.Rows.Count
            pgbReadGremienFromSession.Maximum = SessionGremien.Rows.Count
            My.Settings.AnzSessionGremien = SessionGremien.Rows.Count

            pgbOutlookItems.Maximum = olMandatare.Items.Count
            My.Settings.AnzOutlookItems = olMandatare.Items.Count
            My.Settings.Save()

            If Not RemoveContacts(olMandatare) Then
                Cursor.Current = Cursors.Default
                lblSyncStatus.Text = "ERROR: SYNC FAILED!"
                Exit Sub
            Else
                MakeMandatare(SessionContacts, olMandatare)
                Refresh()
                MakeGremien(SessionGremien, olMandatare)
                Refresh()
            End If
            Cursor.Current = Cursors.Default
            lblSyncStatus.Text = "Die Kontakte wurden erfolgreich syncronisiert!"
            btnClose.Text = "Schließen"
        Catch ex As Exception
            MyError(ex, "RunSync")
            Cursor.Current = Cursors.Default
            lblSyncStatus.Text = ex.Message
            btnClose.Text = "Schließen"
        End Try
    End Sub
    Private Function RemoveContacts(ByRef olMandatare) As Boolean
        Dim x As Integer
        Dim y As Integer
        Dim i As Integer
        Dim anzOutlookItems As Integer
        Try
            anzOutlookItems = olMandatare.Items.Count
            MyLog($"OL deleting {anzOutlookItems} items in folder")

            For i = 1 To anzOutlookItems

                Dim contact As Outlook.ContactItem = TryCast(olMandatare.Items(1), Outlook.ContactItem)
                Dim liste As Outlook.DistListItem = TryCast(olMandatare.Items(1), Outlook.DistListItem)
                If (contact IsNot Nothing) Then
                    x += 1
                    contact.Delete()
                Else
                    y += 1
                    liste.Delete()
                End If
                pgbOutlookItems.Value = anzOutlookItems - i
                lblPgbOutlookItems.Text = anzOutlookItems - i
                Refresh()
                contact = Nothing
                liste = Nothing
            Next
        Catch ex As Exception
            MyError(ex, "OL dele")
            Return False
        End Try
        MyLog("Gelöscht " & (i - 1).ToString & " von " & anzOutlookItems & " K:" & x & "/G:" & y)
        Return True
    End Function
    Sub MakeMandatare(ByRef SessionContacts, ByRef olMandatare)
        Dim cContacts As Integer = 0
        Try
            pgbReadContactsFromSession.Maximum = SessionContacts.rows.count()
            For Each person As DataRow In SessionContacts.Rows
                cContacts += 1

                Dim myDisplayname As String = ""
                If Clean(person("adakgrad")).ToString.Length > 0 Then myDisplayname &= Clean(person("adakgrad")) & " "
                If Clean(person("adtit")).ToString.Length > 0 Then myDisplayname &= Clean(person("adtit")) & " "
                myDisplayname &= Clean(person("advname")) & " " & Clean(person("adname"))
                If Clean(person("adtitn")).ToString.Length > 0 Then myDisplayname &= " " & Clean(person("adtitn"))
                MyLog("--------------------------------------------------------------------------------------------")
                MyLog($"Mandatar: {myDisplayname}")

                '### ToDo change webservices.data.stringlist with dictionary??
                'Dim myCategories As StringList = New StringList From {Clean(person("pepartei"))}
                Dim myCategories As String = Clean(person("pepartei"))

                Dim myFirma As String
                If Clean(person("adfirma1")) = "a" Then myFirma = "" Else myFirma = Clean(person("adfirma1"))

                '### BODY start ###

                Dim myzimmer As String = Clean(person("sbzim"))
                Dim mysbaktiv As String = Clean(person("sbaktiv"))
                Dim mysbatnr As String = Clean(person("sbatnr"))
                Dim mybereich As String = Clean(person("atname"))
                Dim mymodified As String = Clean(person("admodified"))
                Dim myBody As String = "Mitarbeit: " & Clean(person("pepartei")) & vbCrLf
                myBody &= "Sessionnet-Benutzername: " & Clean(person("pesch")) & vbCrLf
                'myBody &= "PRID: " & Clean(person("peid")) & vbCrLf

                If mysbaktiv = "1" Then
                    myBody &= vbCrLf
                    myBody &= "Bearbeiter" & vbCrLf
                    myBody &= "Zimmer: " & myzimmer & vbCrLf
                    myBody &= "Bereich: " & mybereich & vbCrLf
                End If

                myBody &= vbCrLf
                '### ZERODAY - see remarks page down
                Dim zerodate As DateTime = #1/1/1900 01:00:00 AM#

                Dim myadat As String
                Dim myedat As String
                Dim enddate As String = ""
                Dim enddate2 As Date
                Dim startdate As String = ""
                Dim startdate2 As Date
                Dim myperiode As String = ""
                Dim myp As String = "0"

                Dim myGremien As DataTable = GetMyGremien(person("adnr"))
                myBody &= $"Mitglied in {myGremien.Rows.Count} Gremien" & vbCrLf & vbCrLf
                For Each myGremium As DataRow In myGremien.Rows
                    myadat = Clean(myGremium("mgadat"))
                    myedat = Clean(myGremium("mgedat"))
                    Try
                        '### Offsett = 2415019 ; From Session-DB we get a serial 'date'. x - 2415019 you get days since 1.1.1900
                        '### Minimum zeroday is 1.1.1900 01:00 so the number of days is again 2 to much.
                        If myadat > 0 Then
                            startdate = zerodate.AddDays(myadat - 2415019 - 2)
                        Else
                            startdate = ""
                        End If
                        If myedat > 0 Then
                            enddate = zerodate.AddDays(myedat - 2415019 - 2)
                        Else
                            enddate = ""
                        End If
                        If startdate.Length > 0 And enddate.Length > 0 Then
                            startdate2 = startdate
                            enddate2 = enddate
                            myperiode = startdate2.ToString("d") & " bis " & enddate2.ToString("d")
                            myp = 1
                        ElseIf startdate.Length > 0 And enddate.Length = 0 Then
                            startdate2 = startdate
                            myperiode = startdate2.ToString("d") & " bis heute"
                            myp = 2
                        ElseIf startdate.Length = 0 And enddate.Length > 0 Then
                            enddate2 = enddate
                            myperiode = "bis " & enddate2.ToString("d")
                            myp = 3
                        End If
                    Catch ex As Exception
                        MyError(ex, "CalcStartEndDate")
                    Finally
                        MyLog($"gremium: {myGremium("grname")}")
                        MyLog($"myadat: {myadat} myedat: {myedat}")
                        MyLog($"myperiode: {myperiode} {myp}")
                    End Try
                    myBody &= Clean(myGremium("grname")) & " (" & Clean(myGremium("mgfunk")) & ") - " & Clean(myGremium("amname")) & vbCrLf
                    myBody &= vbTab & myperiode & vbCrLf
                    myBody &= vbCrLf
                Next
                myBody &= vbCrLf
                myBody &= "Geändert: " & mymodified & vbCrLf
                '### BODY end ###

                If debug = True Then MyLog($"OL cont New {cContacts} {myDisplayname}")
                Dim newContact As Outlook.ContactItem = olMandatare.Items.Add(Outlook.OlItemType.olContactItem)
                Dim email1 As String = CStr(Clean(person("ademail")))
                Dim email2 As String = CStr(Clean(person("ademail2")))
                If email1.Length > 0 Then newContact.Email1Address = email1
                If email2.Length > 0 Then newContact.Email2Address = email2

                Dim mytitel As String = ""
                Dim adakgrad As String = Clean(person("adakgrad"))
                Dim adtit As String = Clean(person("adtit"))
                Dim adtitn As String = Clean(person("adtitn"))

                If adakgrad.Length > 0 And adtit.Length > 0 Then
                    mytitel = adakgrad & " " & adtit
                ElseIf adakgrad.Length > 0 Then
                    mytitel = adakgrad
                ElseIf adtit.Length > 0 Then
                    mytitel = adtit
                End If

                Dim mygender As String = Clean(person("anname"))
                Dim mygender2 As Outlook.OlGender = Outlook.OlGender.olUnspecified
                If mygender = "Frau" Then
                    mygender2 = Outlook.OlGender.olFemale
                ElseIf mygender = "Herr" Then
                    mygender2 = Outlook.OlGender.olMale
                End If

                With newContact
                    .LastName = Clean(person("adname"))
                    .FirstName = Clean(person("advname"))
                    .CompanyName = myFirma
                    .FullName = myDisplayname
                    .Title = mytitel
                    .Suffix = adtitn
                    .Department = Clean(person("adabt"))
                    .JobTitle = Clean(person("adberuf"))
                    .Profession = Clean(person("adpos"))
                    .Categories = myCategories

                    .HomeTelephoneNumber = Clean(person("adtel2"))
                    .HomeFaxNumber = Clean(person("adtel4"))
                    .Home2TelephoneNumber = Clean(person("adtel6"))

                    .MobileTelephoneNumber = Clean(person("adtel5"))

                    .BusinessAddressCity = Clean(person("adort"))
                    .BusinessAddressCountry = Clean(person("adnat"))
                    .BusinessAddressPostalCode = Clean(person("adplz"))
                    '.BusinessAddressPostOfficeBox
                    '.BusinessAddressState
                    .BusinessAddressStreet = Clean(person("adstr")) & " " & Clean(person("adhnr"))
                    .BusinessTelephoneNumber = Clean(person("adtel"))
                    .BusinessFaxNumber = Clean(person("adtel3"))
                    .BusinessHomePage = Clean(person("adweb"))
                    .Gender = mygender2
                    '.GovernmentIDNumber
                    '.Birthday

                    .Body = myBody

                End With
                newContact.Save()
                pgbReadContactsFromSession.Value = cContacts
                lblPgmSessionContacts.Text = cContacts
                Refresh()
            Next
            My.Settings.AnzOutlookKontakte = cContacts
            My.Settings.Save()
            Refresh()
        Catch ex As Exception
            MyError(ex, "MakeMandatare")
        End Try
    End Sub
    Sub MakeGremien(ByRef SessionGremien As DataTable, ByRef olMandatare As Outlook.MAPIFolder)
        '######
        '### Gremien
        '######
        Try
            Dim cGremien As Integer
            Dim olContactGroupName As String
            Dim outlookObj As Outlook._Application = New Outlook.Application
            Dim contactFolder As Outlook.MAPIFolder = outlookObj.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
            Dim olContactGroupPrefix As String = My.Settings.OlContactGroupPrefix
            My.Settings.AnzOutlookVerteiler = 0

            For Each gremium As DataRow In SessionGremien.Rows
                cGremien += 1
                Dim grNr As Int16 = gremium("grnr")
                Dim grKurz As String = gremium("grkurz").ToString.Trim
                Dim grName As String = gremium("grname").ToString.Trim
                Dim grArt As String = gremium("txname").ToString.Trim
                If grName.Length > 20 Then grName = grKurz
                If grName = "" Then grName = "Kurzname fehlt"
                olContactGroupName = olContactGroupPrefix & " " & grArt & " " & grName
                MyLog($"NEU Verteiler {cGremien} von {SessionGremien.Rows.Count} '{olContactGroupName}'")
                Dim myGremium As DataTable = ReadOneGremiumFromSession(grNr)
                MyLog($"members: {myGremium.Rows.Count}")

                Try
                    Dim newContactGroup As Outlook.DistListItem = Nothing
                    Try
                        newContactGroup = olMandatare.Items.Add(Outlook.OlItemType.olDistributionListItem)
                        newContactGroup.Subject = olContactGroupName
                    Catch ex As Exception
                        MyLog($"OL gr new group: {ex.Message} {olContactGroupName}")
                    End Try
                    Dim newContactGroupErsatz As Outlook.DistListItem = Nothing
                    Try
                        newContactGroupErsatz = olMandatare.Items.Add(Outlook.OlItemType.olDistributionListItem)
                        newContactGroupErsatz.Subject = olContactGroupName & " Ersatzmitglieder"
                    Catch ex As Exception
                        MyLog($"OL gr new groupersatz: {ex.Message} {olContactGroupName}")
                    End Try
                    Dim cNormal As Int16
                    Dim cErsatz As Int16
                    MyLog("adding members now")
                    For Each row As DataRow In myGremium.Rows
                        Dim myName As String
                        myName = row("adname").ToString.Trim & " " & row("advname").ToString.Trim
                        myName &= " [" & row("amname").ToString.Trim & "] " & row("pepartei").ToString.Trim
                        myName &= row("mgfunk").ToString.Trim
                        Dim myEmail As String = row("ademail").ToString.Trim
                        If myEmail.Length > 0 Then
                            Try
                                Dim myDisplayName As String = myName & " (" & myEmail & ")"
                                Dim newrecipient As Outlook.Recipient = outlookObj.Session.CreateRecipient(myDisplayName)
                                newrecipient.Resolve()
                                Select Case row("amname").ToString.Trim
                                    Case "Ordentliches Mitglied", "Zuhörer", "Vorsitz", "Geschäftsführer"
                                        cNormal += 1
                                        newContactGroup.AddMember(newrecipient)
                                    Case "Ersatzmitglied", "Zuhörer Ersatzmitglied"
                                        cErsatz += 1
                                        newContactGroupErsatz.AddMember(newrecipient)
                                End Select
                            Catch ex As Exception
                                MyError(ex, $"OL gr add ERROR: {ex.Message} {olContactGroupName} {myName}")
                            End Try
                        End If
                    Next
                    Try
                        newContactGroup.Save()
                        My.Settings.AnzOutlookVerteiler += 1


                        If cErsatz > 0 Then
                            newContactGroupErsatz.Save()
                            My.Settings.AnzOutlookVerteiler += 1

                        End If
                        pgbReadGremienFromSession.Value = cGremien
                        lblPgmSessionGremien.Text = cGremien
                        Refresh()
                    Catch ex As Exception
                        MyLog($"OL gr new groupe saves: {ex.Message} {olContactGroupName}")
                    End Try
                Catch ex As Exception
                    MyError(ex, $"OL gr new: {ex.Message} {olContactGroupName}")
                    Exit Sub
                End Try
            Next
            My.Settings.Save()
            outlookObj = Nothing
            contactFolder = Nothing
            olContactGroupPrefix = Nothing
        Catch ex As Exception
            MyError(ex, "MakeGremien")
        End Try
    End Sub

    Private Sub FormSync_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Application.DoEvents()
        'RunSync() 'Start direkt bei Formularaufruf
    End Sub

    Private Sub FormSync_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Outlook 2 Session Synchronisation"
        lblSyncStatus.Text = "Bitte starten!"
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Dim button As Button = TryCast(sender, Button)
        Select Case button.Text
            Case "Starten"
                button.Text = "Abbrechen"
                RunSync()
            Case "Abbrechen"
                'Abbruch der Synchronisation ist problematisch, oder?
                'Derzeit nur optischer Aufputz
                button.Text = "Schließen"
            Case "Schließen"
                Me.Close()
        End Select
    End Sub
End Class
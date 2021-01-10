Imports System.Windows.Forms
Imports System.IO.Path
Public Class FormConfig

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

    Public isConfigured As Boolean = False
    Private Sub ConfigForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyLog("### CONFIGFORM-LOAD start")
        Me.KeyPreview = vbTrue
        Me.Text = "Konfiguration - wpContacts Session Outlook Plug-In"
        lblXMLStatus.Text = ""
        tbSQLDataSource.Text = My.Settings.SQLDataSource
        tbSQLInitialCatalog.Text = My.Settings.SQLInitialCatalog
        tbSQLUserID.Text = My.Settings.SQLUserID
        tbSQLPassword.Text = My.Settings.SQLPassword
        tbOlFolder.Text = My.Settings.OlFolder
        tbOlContactGroupPrefix.Text = My.Settings.OlContactGroupPrefix

        AddHandler tbSQLDataSource.TextChanged, AddressOf TbSQLDataSource_TextChanged
        AddHandler tbSQLInitialCatalog.TextChanged, AddressOf TbSQLDataSource_TextChanged
        AddHandler tbOlFolder.TextChanged, AddressOf TbOlFolder_TextChanged

        CheckConfigFormSQL()

        If CheckOutlookFolderExists(My.Settings.OlFolder, "ConfigForm_Load") = False Then
            BtnOlNew.Enabled = True
            tbOlFolder.ForeColor = Drawing.Color.Red
        End If
        MyLog("### CONFIGFORM-LOAD end")
    End Sub

    Public Function CheckConfigFormSQL() As Boolean
        If tbSQLDataSource.Text = "" Or tbSQLInitialCatalog.Text = "" Then
            BtnSQLtest.Enabled = False
            MyLog("CheckConfigFormSQL: SQL Connection credentials empty")
            lblSQLStatus.Text = "SQL Server Verbindung: Fehler"
            lblSQLStatus.ForeColor = Drawing.Color.Red
        Else
            BtnSQLtest.Enabled = True
            If SqlTestConnection(tbSQLDataSource.Text, tbSQLInitialCatalog.Text, tbSQLUserID.Text, tbSQLPassword.Text) = True Then
                MyLog("CheckConfigFormSQL: SQL Connection OK")
                lblSQLStatus.Text = "SQL Server Verbindung: OK"
                lblSQLStatus.ForeColor = Drawing.Color.Green
            Else
                MyLog("CheckConfigFormSQL: SQL Connection failed", "error")
                lblSQLStatus.Text = "SQL Server Verbindung: Fehler"
                lblSQLStatus.ForeColor = Drawing.Color.Red
            End If
        End If
        Return True
    End Function
    Private Sub ConfigForm_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub TbSQLDataSource_TextChanged(sender As Object, e As EventArgs) 'Handles tbSQLDataSource.TextChanged, tbSQLInitialCatalog.TextChanged
        If tbSQLDataSource.Text = "" Or tbSQLInitialCatalog.Text = "" Then
            BtnSQLtest.Enabled = False
        Else
            BtnSQLtest.Enabled = True
        End If
    End Sub
    Private Sub TbOlFolder_TextChanged(sender As Object, e As EventArgs) 'Handles tbOlFolder.TextChanged
        If CheckOutlookFolderExists(tbOlFolder.Text, "TbOlFolder_TextChanged") = False Then
            BtnOlNew.Enabled = True
            tbOlFolder.ForeColor = Drawing.Color.Red
        Else
            BtnOlNew.Enabled = False
            tbOlFolder.ForeColor = Drawing.Color.Green

        End If
    End Sub

    Private Sub BtnXMLload_Click(sender As Object, e As EventArgs) Handles BtnXMLload.Click
        Try
            MyLog("Loading XML")
            MyLog("Old logfile: " & My.Settings.LastConfigFile)
            lblXMLStatus.Text = ""
            OpenFileDialog1.RestoreDirectory = True
            OpenFileDialog1.ShowHelp = True
            OpenFileDialog1.Title = "XML Konfiguration öffnen"
            OpenFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            If My.Settings.LastConfigFile = "" Then
                OpenFileDialog1.InitialDirectory = "c:\"
                OpenFileDialog1.FileName = "config.xml"
            Else
                OpenFileDialog1.InitialDirectory = GetDirectoryName(My.Settings.LastConfigFile)
                OpenFileDialog1.FileName = GetFileName(My.Settings.LastConfigFile)
            End If

            Dim result As DialogResult = OpenFileDialog1.ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                My.Settings.LastConfigFile = OpenFileDialog1.FileName
                My.Settings.Save()
                MyLog("Selected XML is: " & OpenFileDialog1.FileName)
                Dim config As Dictionary(Of String, String) = ReadConfigFromXML(OpenFileDialog1.FileName)
                tbSQLDataSource.Text = config("SQLDataSource")
                tbSQLInitialCatalog.Text = config("SQLInitialCatalog")
                tbSQLUserID.Text = config("SQLUserID")
                tbSQLPassword.Text = config("SQLPassword")
                tbOlContactGroupPrefix.Text = config("olContactGroupPrefix")
                tbOlFolder.Text = config("olFolder")
                lblXMLStatus.Text = "XML geladen"
            Else
                MyLog("no XML file selected")
            End If
            CheckConfigFormSQL()
        Catch ex As Exception
            MyLog(ex.Message)
        End Try
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        Dim sqlok As Boolean
        Dim outlookok As Boolean
        My.Settings.SQLDataSource = tbSQLDataSource.Text
        My.Settings.SQLInitialCatalog = tbSQLInitialCatalog.Text
        My.Settings.SQLUserID = tbSQLUserID.Text
        My.Settings.SQLPassword = tbSQLPassword.Text
        My.Settings.OlContactGroupPrefix = tbOlContactGroupPrefix.Text
        My.Settings.OlFolder = tbOlFolder.Text
        My.Settings.isConfigured = False
        My.Settings.Save()
        sqlok = SqlTestConnection(My.Settings.SQLDataSource, My.Settings.SQLInitialCatalog, My.Settings.SQLUserID, My.Settings.SQLPassword)
        outlookok = CheckOutlookFolderExists(My.Settings.OlFolder)
        If sqlok = True And outlookok = True Then
            My.Settings.isConfigured = True
            My.Settings.Save()
        End If
        Close()
    End Sub

    Private Sub BtnXMLsave_Click(sender As Object, e As EventArgs) Handles BtnXMLsave.Click
        MyLog("### SavingXML start")
        Try

            MyLog("BtnXMLsave_Click: last configfile = " & My.Settings.LastConfigFile)
            lblXMLStatus.Text = ""
            SaveFileDialog1.RestoreDirectory = True
            SaveFileDialog1.ShowHelp = True
            SaveFileDialog1.Title = "XML Konfiguration öffnen"
            SaveFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            If My.Settings.LastConfigFile = "" Then
                SaveFileDialog1.FileName = "config.xml"
            Else
                SaveFileDialog1.InitialDirectory = GetDirectoryName(My.Settings.LastConfigFile)
                SaveFileDialog1.FileName = GetFileName(My.Settings.LastConfigFile)
            End If

            Dim result As DialogResult = SaveFileDialog1.ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                My.Settings.LastConfigFile = SaveFileDialog1.FileName
                My.Settings.Save()
                SaveConfigToXML(SaveFileDialog1.FileName)
                lblXMLStatus.Text = "XML Konfiguration gespeichert!"
            End If

        Catch ex As Exception
            MyLog($"BtnXMLsave_Click: {ex.Message}", "error")
        End Try
        MyLog("### SavingXML end")
    End Sub

    Private Sub BtSQLtest_Click(sender As Object, e As EventArgs) Handles BtnSQLtest.Click
        If SqlTestConnection(tbSQLDataSource.Text, tbSQLInitialCatalog.Text, tbSQLUserID.Text, tbSQLPassword.Text) = True Then
            MyLog("SQL Connection OK")
            lblSQLStatus.Text = "SQL Server Verbindung: OK"
            lblSQLStatus.ForeColor = Drawing.Color.Green
        Else
            MyLog("SQL Connection failed")
            lblSQLStatus.Text = "SQL Server Verbindung: Fehler"
            lblSQLStatus.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub BtnOlNew_Click(sender As Object, e As EventArgs) Handles BtnOlNew.Click
        Try
            If My.Settings.OlFolder <> tbOlFolder.Text Then
                My.Settings.OlFolder = tbOlFolder.Text
                My.Settings.Save()
            End If
            MyLog("Creating folder " & My.Settings.OlFolder)
            If CreateOutlookFolder(My.Settings.OlFolder) = True Then
                BtnOlNew.Enabled = False
                tbOlFolder.ForeColor = Drawing.SystemColors.WindowText
            End If
        Catch ex As Exception
            MyLog(ex.Message)
        End Try

    End Sub


    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

End Class
﻿Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms
Imports System.Deployment.Application

'### https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.contactitem?view=outlook-pia
'### https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.distlistitem?view=outlook-pia

Public Class WpContactsRibbon1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim sqlok As Boolean = False
        Dim outlookok As Boolean = False
        MyLog("### RIBBONLOAD start")
        My.Settings.isConfigured = False
        SetNetworkConfigPath()
        sqlok = SqlTestConnection(My.Settings.SQLDataSource, My.Settings.SQLInitialCatalog, My.Settings.SQLUserID, My.Settings.SQLPassword)
        outlookok = CheckOutlookFolderExists(My.Settings.OlFolder)
        If sqlok = True And outlookok = True Then My.Settings.isConfigured = True
        My.Settings.Save()
        ListSettings()

        If My.Settings.isConfigured = True Then
            BtnUpdate.Enabled = True
            AnzSessionMitarbeiter.Text = My.Settings.AnzSessionMitarbeiter
            AnzSessionGremien.Text = My.Settings.AnzSessionGremien
            AnzOutlookItems.Text = My.Settings.AnzOutlookItems
            AnzOutlookVerteiler.Text = My.Settings.AnzOutlookVerteiler
            AnzOutlookKontakte.Text = My.Settings.AnzOutlookKontakte
            lblKontakteordner.Label = My.Settings.OlFolder
        Else
            LoadNetworkConfiguration()
            CreateOutlookFolder(My.Settings.OlFolder)
            BtnUpdate.Enabled = True
        End If

        MyLog("### RIBBONLOAD end")
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnUpdate.Click
        '### VERSION CHECK
        MyLog("----------------------------------------------------------------------------------")
        Dim versionNumber As String '= Assembly.GetExecutingAssembly().GetName().Version.ToString
        If ApplicationDeployment.IsNetworkDeployed Then
            Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
            versionNumber = "v " & AD.CurrentVersion.ToString
        Else
            versionNumber = "v " & My.Application.Info.Version.ToString
        End If
        MyLog($"wpContact-Outlook-Add-in ({versionNumber}) startet!")

        Dim f As New FormSync
        f.ShowDialog()
        f.Dispose()
        AnzSessionMitarbeiter.Text = My.Settings.AnzSessionMitarbeiter
        AnzSessionGremien.Text = My.Settings.AnzSessionGremien
        AnzOutlookItems.Text = My.Settings.AnzOutlookItems
        AnzOutlookVerteiler.Text = My.Settings.AnzOutlookVerteiler
        AnzOutlookKontakte.Text = My.Settings.AnzOutlookKontakte
        lblKontakteordner.Label = My.Settings.OlFolder
    End Sub

    'Private Sub BtnCheckUPD_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnCheckUPD.Click
    '    InstallUpdateSyncWithInfo()
    'End Sub

    Private Sub BtnConfig_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnConfig.Click
        MyLog("### BtnConfig_Click start")
        Dim f = New FormConfig
        f.ShowDialog()
        If My.Settings.isConfigured = True Then
            BtnUpdate.Enabled = True
        Else
            BtnUpdate.Enabled = False
        End If
        MyLog("### BtnConfig_Click end")
    End Sub

    Private Sub BtnInfo_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnInfo.Click
        Dim f = New AboutBox1
        f.ShowDialog()
    End Sub

    Private Sub BtnNetworkConfig_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnNetworkConfig.Click
        LoadNetworkConfiguration()
    End Sub
    Private Function LoadNetworkConfiguration()
        Dim configfile As String = My.Settings.NetworkConfigStore & "\" & My.Settings.DefaultConfigFile
        Try

            If IO.File.Exists(configfile) Then

                Dim config As Dictionary(Of String, String) = ReadConfigFromXML(configfile)
                My.Settings.SQLDataSource = config("SQLDataSource")
                My.Settings.SQLInitialCatalog = config("SQLInitialCatalog")
                My.Settings.SQLUserID = config("SQLUserID")
                My.Settings.SQLPassword = config("SQLPassword")
                My.Settings.OlContactGroupPrefix = config("olContactGroupPrefix")
                My.Settings.OlFolder = config("olFolder")
            Else
                MyLog($"NetworkConfig not found: {configfile}")
            End If
            Return True
        Catch ex As Exception
            MyLog($"Error LoadNetworkConfiguration: {ex.Message} {configfile}", "error")
            Return False
        End Try
    End Function

    Private Sub BtnResetSettings_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnResetSettings.Click
        My.Settings.Reset()
    End Sub
End Class
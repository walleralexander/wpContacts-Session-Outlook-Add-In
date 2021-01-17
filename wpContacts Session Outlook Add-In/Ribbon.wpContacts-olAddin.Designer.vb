Partial Class WpContactsRibbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()

    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    'Das Bearbeiten ist mit dem Komponenten-Designer möglich.
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.BefehleGroup = Me.Factory.CreateRibbonGroup
        Me.SessionGroup2 = Me.Factory.CreateRibbonGroup
        Me.AnzSessionMitarbeiter = Me.Factory.CreateRibbonEditBox
        Me.AnzSessionGremien = Me.Factory.CreateRibbonEditBox
        Me.OutlookGroup = Me.Factory.CreateRibbonGroup
        Me.AnzOutlookKontakte = Me.Factory.CreateRibbonEditBox
        Me.AnzOutlookVerteiler = Me.Factory.CreateRibbonEditBox
        Me.AnzOutlookItems = Me.Factory.CreateRibbonEditBox
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.lblKontakteordner = Me.Factory.CreateRibbonLabel
        Me.lblDebug = Me.Factory.CreateRibbonLabel
        Me.BtnUpdate = Me.Factory.CreateRibbonButton
        Me.BtnGremien = Me.Factory.CreateRibbonButton
        Me.BtnConfig = Me.Factory.CreateRibbonButton
        Me.BtnNetworkConfig = Me.Factory.CreateRibbonButton
        Me.BtnResetSettings = Me.Factory.CreateRibbonButton
        Me.BtnInfo = Me.Factory.CreateRibbonButton
        Me.BtnCheckUPD = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.BefehleGroup.SuspendLayout()
        Me.SessionGroup2.SuspendLayout()
        Me.OutlookGroup.SuspendLayout()
        Me.Box1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.BefehleGroup)
        Me.Tab1.Groups.Add(Me.SessionGroup2)
        Me.Tab1.Groups.Add(Me.OutlookGroup)
        Me.Tab1.Label = "wpContacts"
        Me.Tab1.Name = "Tab1"
        '
        'BefehleGroup
        '
        Me.BefehleGroup.Items.Add(Me.BtnUpdate)
        Me.BefehleGroup.Items.Add(Me.BtnGremien)
        Me.BefehleGroup.Items.Add(Me.BtnConfig)
        Me.BefehleGroup.Items.Add(Me.BtnInfo)
        Me.BefehleGroup.Items.Add(Me.BtnCheckUPD)
        Me.BefehleGroup.Items.Add(Me.BtnNetworkConfig)
        Me.BefehleGroup.Items.Add(Me.BtnResetSettings)
        Me.BefehleGroup.Label = "Befehle"
        Me.BefehleGroup.Name = "BefehleGroup"
        '
        'SessionGroup2
        '
        Me.SessionGroup2.Items.Add(Me.AnzSessionMitarbeiter)
        Me.SessionGroup2.Items.Add(Me.AnzSessionGremien)
        Me.SessionGroup2.Label = "Session"
        Me.SessionGroup2.Name = "SessionGroup2"
        '
        'AnzSessionMitarbeiter
        '
        Me.AnzSessionMitarbeiter.Enabled = False
        Me.AnzSessionMitarbeiter.Label = "Mitarbeiter"
        Me.AnzSessionMitarbeiter.Name = "AnzSessionMitarbeiter"
        Me.AnzSessionMitarbeiter.Text = Nothing
        '
        'AnzSessionGremien
        '
        Me.AnzSessionGremien.Enabled = False
        Me.AnzSessionGremien.Label = "Gremien"
        Me.AnzSessionGremien.Name = "AnzSessionGremien"
        Me.AnzSessionGremien.Text = Nothing
        '
        'OutlookGroup
        '
        Me.OutlookGroup.Items.Add(Me.AnzOutlookKontakte)
        Me.OutlookGroup.Items.Add(Me.AnzOutlookVerteiler)
        Me.OutlookGroup.Items.Add(Me.AnzOutlookItems)
        Me.OutlookGroup.Items.Add(Me.Separator1)
        Me.OutlookGroup.Items.Add(Me.Box1)
        Me.OutlookGroup.Label = "Outlook"
        Me.OutlookGroup.Name = "OutlookGroup"
        '
        'AnzOutlookKontakte
        '
        Me.AnzOutlookKontakte.Enabled = False
        Me.AnzOutlookKontakte.Label = "Kontakte"
        Me.AnzOutlookKontakte.Name = "AnzOutlookKontakte"
        Me.AnzOutlookKontakte.SuperTip = "Anzahl der Kontakte im Outlook (berechnet)"
        Me.AnzOutlookKontakte.Text = Nothing
        '
        'AnzOutlookVerteiler
        '
        Me.AnzOutlookVerteiler.Enabled = False
        Me.AnzOutlookVerteiler.Label = "Verteiler"
        Me.AnzOutlookVerteiler.Name = "AnzOutlookVerteiler"
        Me.AnzOutlookVerteiler.SuperTip = "Anzahl der Verteiler inklusive der Verteiler für Ersatzmitglieder."
        Me.AnzOutlookVerteiler.Text = Nothing
        '
        'AnzOutlookItems
        '
        Me.AnzOutlookItems.Enabled = False
        Me.AnzOutlookItems.Label = "Kontakte && Verteiler"
        Me.AnzOutlookItems.Name = "AnzOutlookItems"
        Me.AnzOutlookItems.SuperTip = "Eintragungen im Kontakteordner"
        Me.AnzOutlookItems.Text = Nothing
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Box1
        '
        Me.Box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box1.Items.Add(Me.Label1)
        Me.Box1.Items.Add(Me.lblKontakteordner)
        Me.Box1.Items.Add(Me.lblDebug)
        Me.Box1.Name = "Box1"
        '
        'Label1
        '
        Me.Label1.Label = "Ordner:"
        Me.Label1.Name = "Label1"
        '
        'lblKontakteordner
        '
        Me.lblKontakteordner.Label = "lblKontakteordner"
        Me.lblKontakteordner.Name = "lblKontakteordner"
        '
        'lblDebug
        '
        Me.lblDebug.Label = "[DEBUGMODE]"
        Me.lblDebug.Name = "lblDebug"
        Me.lblDebug.Visible = False
        '
        'BtnUpdate
        '
        Me.BtnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnUpdate.Description = "Session Kontakte aktualisieren"
        Me.BtnUpdate.Enabled = False
        Me.BtnUpdate.Label = "Synchronisation"
        Me.BtnUpdate.Name = "BtnUpdate"
        Me.BtnUpdate.OfficeImageId = "Zoom100"
        Me.BtnUpdate.ShowImage = True
        '
        'BtnGremien
        '
        Me.BtnGremien.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnGremien.Image = Global.wpContact_Outlook_Add_in.My.Resources.Resources.ListViewTable_16x
        Me.BtnGremien.Label = "Gremien"
        Me.BtnGremien.Name = "BtnGremien"
        Me.BtnGremien.ShowImage = True
        '
        'BtnConfig
        '
        Me.BtnConfig.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnConfig.Label = "Konfiguration"
        Me.BtnConfig.Name = "BtnConfig"
        Me.BtnConfig.OfficeImageId = "AddInManager"
        Me.BtnConfig.ShowImage = True
        '
        'BtnNetworkConfig
        '
        Me.BtnNetworkConfig.Label = "Test NetworkConfig"
        Me.BtnNetworkConfig.Name = "BtnNetworkConfig"
        Me.BtnNetworkConfig.Visible = False
        '
        'BtnResetSettings
        '
        Me.BtnResetSettings.Label = "Reset Settings"
        Me.BtnResetSettings.Name = "BtnResetSettings"
        Me.BtnResetSettings.Visible = False
        '
        'BtnInfo
        '
        Me.BtnInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnInfo.Image = Global.wpContact_Outlook_Add_in.My.Resources.Resources.StatusInformation_16x
        Me.BtnInfo.Label = "Info"
        Me.BtnInfo.Name = "BtnInfo"
        Me.BtnInfo.ShowImage = True
        '
        'BtnCheckUPD
        '
        Me.BtnCheckUPD.Enabled = False
        Me.BtnCheckUPD.Image = Global.wpContact_Outlook_Add_in.My.Resources.Resources.GetLatestVersion_16x
        Me.BtnCheckUPD.Label = "Check Version"
        Me.BtnCheckUPD.Name = "BtnCheckUPD"
        Me.BtnCheckUPD.ShowImage = True
        Me.BtnCheckUPD.Visible = False
        '
        'WpContactsRibbon1
        '
        Me.Name = "WpContactsRibbon1"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.BefehleGroup.ResumeLayout(False)
        Me.BefehleGroup.PerformLayout()
        Me.SessionGroup2.ResumeLayout(False)
        Me.SessionGroup2.PerformLayout()
        Me.OutlookGroup.ResumeLayout(False)
        Me.OutlookGroup.PerformLayout()
        Me.Box1.ResumeLayout(False)
        Me.Box1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents BefehleGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnUpdate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AnzOutlookItems As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents SessionGroup2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AnzSessionMitarbeiter As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents OutlookGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AnzSessionGremien As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents BtnConfig As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnCheckUPD As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnInfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AnzOutlookVerteiler As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents AnzOutlookKontakte As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents lblKontakteordner As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Box1 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents BtnNetworkConfig As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnResetSettings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents lblDebug As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents BtnGremien As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As wpContactsRibbon1
        Get
            Return Me.GetRibbon(Of wpContactsRibbon1)()
        End Get
    End Property
End Class

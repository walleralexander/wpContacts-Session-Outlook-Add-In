<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormSync
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.pgbOutlookItems = New System.Windows.Forms.ProgressBar()
        Me.pgbReadGremienFromSession = New System.Windows.Forms.ProgressBar()
        Me.pgbReadContactsFromSession = New System.Windows.Forms.ProgressBar()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.lblSyncStatus = New System.Windows.Forms.Label()
        Me.cbxSQLConnection = New System.Windows.Forms.CheckBox()
        Me.cbxOutlook = New System.Windows.Forms.CheckBox()
        Me.cbxFolder = New System.Windows.Forms.CheckBox()
        Me.cbxReadContactsFromSession = New System.Windows.Forms.CheckBox()
        Me.cbxReadGremienFromSession = New System.Windows.Forms.CheckBox()
        Me.lblReadContactsFromSession = New System.Windows.Forms.Label()
        Me.lblReadGremienFromSession = New System.Windows.Forms.Label()
        Me.lblOutlookItems = New System.Windows.Forms.Label()
        Me.lblPgbOutlookItems = New System.Windows.Forms.Label()
        Me.lblPgmSessionContacts = New System.Windows.Forms.Label()
        Me.lblPgmSessionGremien = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'pgbOutlookItems
        '
        Me.pgbOutlookItems.BackColor = System.Drawing.SystemColors.Control
        Me.pgbOutlookItems.Location = New System.Drawing.Point(18, 228)
        Me.pgbOutlookItems.Name = "pgbOutlookItems"
        Me.pgbOutlookItems.Size = New System.Drawing.Size(500, 40)
        Me.pgbOutlookItems.TabIndex = 0
        '
        'pgbReadGremienFromSession
        '
        Me.pgbReadGremienFromSession.Location = New System.Drawing.Point(18, 374)
        Me.pgbReadGremienFromSession.Name = "pgbReadGremienFromSession"
        Me.pgbReadGremienFromSession.Size = New System.Drawing.Size(500, 40)
        Me.pgbReadGremienFromSession.TabIndex = 1
        '
        'pgbReadContactsFromSession
        '
        Me.pgbReadContactsFromSession.Location = New System.Drawing.Point(18, 301)
        Me.pgbReadContactsFromSession.Name = "pgbReadContactsFromSession"
        Me.pgbReadContactsFromSession.Size = New System.Drawing.Size(500, 40)
        Me.pgbReadContactsFromSession.TabIndex = 2
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(195, 427)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(137, 36)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Starten"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'lblSyncStatus
        '
        Me.lblSyncStatus.Location = New System.Drawing.Point(12, 9)
        Me.lblSyncStatus.Name = "lblSyncStatus"
        Me.lblSyncStatus.Size = New System.Drawing.Size(504, 36)
        Me.lblSyncStatus.TabIndex = 7
        Me.lblSyncStatus.Text = "Status"
        '
        'cbxSQLConnection
        '
        Me.cbxSQLConnection.AutoSize = True
        Me.cbxSQLConnection.Location = New System.Drawing.Point(18, 109)
        Me.cbxSQLConnection.Name = "cbxSQLConnection"
        Me.cbxSQLConnection.Size = New System.Drawing.Size(203, 24)
        Me.cbxSQLConnection.TabIndex = 8
        Me.cbxSQLConnection.Text = "Verbindung zu Session"
        Me.cbxSQLConnection.UseVisualStyleBackColor = True
        '
        'cbxOutlook
        '
        Me.cbxOutlook.AutoSize = True
        Me.cbxOutlook.Location = New System.Drawing.Point(18, 79)
        Me.cbxOutlook.Name = "cbxOutlook"
        Me.cbxOutlook.Size = New System.Drawing.Size(200, 24)
        Me.cbxOutlook.TabIndex = 9
        Me.cbxOutlook.Text = "Verbindung zu Outlook"
        Me.cbxOutlook.UseVisualStyleBackColor = True
        '
        'cbxFolder
        '
        Me.cbxFolder.AutoSize = True
        Me.cbxFolder.Location = New System.Drawing.Point(18, 49)
        Me.cbxFolder.Name = "cbxFolder"
        Me.cbxFolder.Size = New System.Drawing.Size(140, 24)
        Me.cbxFolder.TabIndex = 10
        Me.cbxFolder.Text = "Suche Ordner:"
        Me.cbxFolder.UseVisualStyleBackColor = True
        '
        'cbxReadContactsFromSession
        '
        Me.cbxReadContactsFromSession.AutoSize = True
        Me.cbxReadContactsFromSession.Location = New System.Drawing.Point(18, 140)
        Me.cbxReadContactsFromSession.Name = "cbxReadContactsFromSession"
        Me.cbxReadContactsFromSession.Size = New System.Drawing.Size(194, 24)
        Me.cbxReadContactsFromSession.TabIndex = 11
        Me.cbxReadContactsFromSession.Text = "Einlesen der Kontakte"
        Me.cbxReadContactsFromSession.UseVisualStyleBackColor = True
        '
        'cbxReadGremienFromSession
        '
        Me.cbxReadGremienFromSession.AutoSize = True
        Me.cbxReadGremienFromSession.Location = New System.Drawing.Point(18, 171)
        Me.cbxReadGremienFromSession.Name = "cbxReadGremienFromSession"
        Me.cbxReadGremienFromSession.Size = New System.Drawing.Size(193, 24)
        Me.cbxReadGremienFromSession.TabIndex = 12
        Me.cbxReadGremienFromSession.Text = "Einlesen der Gremien"
        Me.cbxReadGremienFromSession.UseVisualStyleBackColor = True
        '
        'lblReadContactsFromSession
        '
        Me.lblReadContactsFromSession.Location = New System.Drawing.Point(287, 138)
        Me.lblReadContactsFromSession.Name = "lblReadContactsFromSession"
        Me.lblReadContactsFromSession.Size = New System.Drawing.Size(45, 24)
        Me.lblReadContactsFromSession.TabIndex = 13
        Me.lblReadContactsFromSession.Text = "0"
        Me.lblReadContactsFromSession.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblReadGremienFromSession
        '
        Me.lblReadGremienFromSession.Location = New System.Drawing.Point(287, 169)
        Me.lblReadGremienFromSession.Name = "lblReadGremienFromSession"
        Me.lblReadGremienFromSession.Size = New System.Drawing.Size(45, 24)
        Me.lblReadGremienFromSession.TabIndex = 14
        Me.lblReadGremienFromSession.Text = "0"
        Me.lblReadGremienFromSession.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblOutlookItems
        '
        Me.lblOutlookItems.Location = New System.Drawing.Point(287, 77)
        Me.lblOutlookItems.Name = "lblOutlookItems"
        Me.lblOutlookItems.Size = New System.Drawing.Size(45, 24)
        Me.lblOutlookItems.TabIndex = 15
        Me.lblOutlookItems.Text = "0"
        Me.lblOutlookItems.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblPgbOutlookItems
        '
        Me.lblPgbOutlookItems.Location = New System.Drawing.Point(473, 200)
        Me.lblPgbOutlookItems.Name = "lblPgbOutlookItems"
        Me.lblPgbOutlookItems.Size = New System.Drawing.Size(45, 20)
        Me.lblPgbOutlookItems.TabIndex = 16
        Me.lblPgbOutlookItems.Text = "0"
        Me.lblPgbOutlookItems.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPgmSessionContacts
        '
        Me.lblPgmSessionContacts.Location = New System.Drawing.Point(473, 275)
        Me.lblPgmSessionContacts.Name = "lblPgmSessionContacts"
        Me.lblPgmSessionContacts.Size = New System.Drawing.Size(45, 20)
        Me.lblPgmSessionContacts.TabIndex = 17
        Me.lblPgmSessionContacts.Text = "0"
        Me.lblPgmSessionContacts.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPgmSessionGremien
        '
        Me.lblPgmSessionGremien.Location = New System.Drawing.Point(473, 350)
        Me.lblPgmSessionGremien.Name = "lblPgmSessionGremien"
        Me.lblPgmSessionGremien.Size = New System.Drawing.Size(45, 20)
        Me.lblPgmSessionGremien.TabIndex = 18
        Me.lblPgmSessionGremien.Text = "0"
        Me.lblPgmSessionGremien.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 204)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(234, 20)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Lösche Kontakte und Verteiler"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 277)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(141, 20)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Erzeuge Kontakte"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 350)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(139, 20)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Erzeuge Verteiler"
        '
        'FormSync
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(538, 496)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblPgmSessionGremien)
        Me.Controls.Add(Me.lblPgmSessionContacts)
        Me.Controls.Add(Me.lblPgbOutlookItems)
        Me.Controls.Add(Me.lblOutlookItems)
        Me.Controls.Add(Me.lblReadGremienFromSession)
        Me.Controls.Add(Me.lblReadContactsFromSession)
        Me.Controls.Add(Me.cbxReadGremienFromSession)
        Me.Controls.Add(Me.cbxReadContactsFromSession)
        Me.Controls.Add(Me.cbxFolder)
        Me.Controls.Add(Me.cbxOutlook)
        Me.Controls.Add(Me.cbxSQLConnection)
        Me.Controls.Add(Me.lblSyncStatus)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.pgbReadContactsFromSession)
        Me.Controls.Add(Me.pgbReadGremienFromSession)
        Me.Controls.Add(Me.pgbOutlookItems)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormSync"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form.Sync"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents pgbOutlookItems As Windows.Forms.ProgressBar
    Friend WithEvents pgbReadGremienFromSession As Windows.Forms.ProgressBar
    Friend WithEvents pgbReadContactsFromSession As Windows.Forms.ProgressBar
    Friend WithEvents btnClose As Windows.Forms.Button
    Friend WithEvents lblSyncStatus As Windows.Forms.Label
    Friend WithEvents cbxSQLConnection As Windows.Forms.CheckBox
    Friend WithEvents cbxOutlook As Windows.Forms.CheckBox
    Friend WithEvents cbxFolder As Windows.Forms.CheckBox
    Friend WithEvents cbxReadContactsFromSession As Windows.Forms.CheckBox
    Friend WithEvents cbxReadGremienFromSession As Windows.Forms.CheckBox
    Friend WithEvents lblReadContactsFromSession As Windows.Forms.Label
    Friend WithEvents lblReadGremienFromSession As Windows.Forms.Label
    Friend WithEvents lblOutlookItems As Windows.Forms.Label
    Friend WithEvents lblPgbOutlookItems As Windows.Forms.Label
    Friend WithEvents lblPgmSessionContacts As Windows.Forms.Label
    Friend WithEvents lblPgmSessionGremien As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
End Class

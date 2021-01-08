<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormConfig
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblSQLStatus = New System.Windows.Forms.Label()
        Me.BtnSQLtest = New System.Windows.Forms.Button()
        Me.tbSQLPassword = New System.Windows.Forms.TextBox()
        Me.tbSQLUserID = New System.Windows.Forms.TextBox()
        Me.tbSQLInitialCatalog = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbSQLDataSource = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.BtnOlNew = New System.Windows.Forms.Button()
        Me.tbOlContactGroupPrefix = New System.Windows.Forms.TextBox()
        Me.tbOlFolder = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.BtnXMLload = New System.Windows.Forms.Button()
        Me.BtnXMLsave = New System.Windows.Forms.Button()
        Me.BtnSave = New System.Windows.Forms.Button()
        Me.lblXMLStatus = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnCancel
        '
        Me.BtnCancel.Location = New System.Drawing.Point(304, 278)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(140, 23)
        Me.BtnCancel.TabIndex = 0
        Me.BtnCancel.Text = "Abbruch"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblSQLStatus)
        Me.GroupBox1.Controls.Add(Me.BtnSQLtest)
        Me.GroupBox1.Controls.Add(Me.tbSQLPassword)
        Me.GroupBox1.Controls.Add(Me.tbSQLUserID)
        Me.GroupBox1.Controls.Add(Me.tbSQLInitialCatalog)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.tbSQLDataSource)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(432, 169)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Session Anbindung"
        '
        'lblSQLStatus
        '
        Me.lblSQLStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSQLStatus.Location = New System.Drawing.Point(6, 121)
        Me.lblSQLStatus.Name = "lblSQLStatus"
        Me.lblSQLStatus.Size = New System.Drawing.Size(286, 26)
        Me.lblSQLStatus.TabIndex = 7
        Me.lblSQLStatus.Text = "leer"
        Me.lblSQLStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BtnSQLtest
        '
        Me.BtnSQLtest.Enabled = False
        Me.BtnSQLtest.Location = New System.Drawing.Point(307, 124)
        Me.BtnSQLtest.Name = "BtnSQLtest"
        Me.BtnSQLtest.Size = New System.Drawing.Size(98, 23)
        Me.BtnSQLtest.TabIndex = 6
        Me.BtnSQLtest.Text = "SQL testen"
        Me.BtnSQLtest.UseVisualStyleBackColor = True
        '
        'tbSQLPassword
        '
        Me.tbSQLPassword.Location = New System.Drawing.Point(155, 98)
        Me.tbSQLPassword.Name = "tbSQLPassword"
        Me.tbSQLPassword.Size = New System.Drawing.Size(250, 20)
        Me.tbSQLPassword.TabIndex = 4
        '
        'tbSQLUserID
        '
        Me.tbSQLUserID.Location = New System.Drawing.Point(155, 71)
        Me.tbSQLUserID.Name = "tbSQLUserID"
        Me.tbSQLUserID.Size = New System.Drawing.Size(250, 20)
        Me.tbSQLUserID.TabIndex = 3
        '
        'tbSQLInitialCatalog
        '
        Me.tbSQLInitialCatalog.Location = New System.Drawing.Point(155, 44)
        Me.tbSQLInitialCatalog.Name = "tbSQLInitialCatalog"
        Me.tbSQLInitialCatalog.Size = New System.Drawing.Size(250, 20)
        Me.tbSQLInitialCatalog.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 101)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(84, 15)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "SQL Passwort"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 74)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 15)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "SQL Benutzername"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "SQL Datenbank"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "SQL Datenquelle"
        '
        'tbSQLDataSource
        '
        Me.tbSQLDataSource.Location = New System.Drawing.Point(155, 17)
        Me.tbSQLDataSource.Name = "tbSQLDataSource"
        Me.tbSQLDataSource.Size = New System.Drawing.Size(250, 20)
        Me.tbSQLDataSource.TabIndex = 0
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnOlNew)
        Me.GroupBox2.Controls.Add(Me.tbOlContactGroupPrefix)
        Me.GroupBox2.Controls.Add(Me.tbOlFolder)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 187)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(432, 85)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Outlook"
        '
        'BtnOlNew
        '
        Me.BtnOlNew.Enabled = False
        Me.BtnOlNew.Location = New System.Drawing.Point(362, 15)
        Me.BtnOlNew.Name = "BtnOlNew"
        Me.BtnOlNew.Size = New System.Drawing.Size(43, 23)
        Me.BtnOlNew.TabIndex = 9
        Me.BtnOlNew.Text = "NEU"
        Me.BtnOlNew.UseVisualStyleBackColor = True
        '
        'tbOlContactGroupPrefix
        '
        Me.tbOlContactGroupPrefix.Location = New System.Drawing.Point(155, 46)
        Me.tbOlContactGroupPrefix.Name = "tbOlContactGroupPrefix"
        Me.tbOlContactGroupPrefix.Size = New System.Drawing.Size(200, 20)
        Me.tbOlContactGroupPrefix.TabIndex = 8
        '
        'tbOlFolder
        '
        Me.tbOlFolder.Location = New System.Drawing.Point(155, 19)
        Me.tbOlFolder.Name = "tbOlFolder"
        Me.tbOlFolder.Size = New System.Drawing.Size(200, 20)
        Me.tbOlFolder.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 49)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(85, 15)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Gruppenprefix"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(84, 15)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Kontaktordner"
        '
        'BtnXMLload
        '
        Me.BtnXMLload.Image = Global.wpContact_Outlook_Add_in.My.Resources.Resources.SQLDatabase_16x
        Me.BtnXMLload.Location = New System.Drawing.Point(12, 278)
        Me.BtnXMLload.Name = "BtnXMLload"
        Me.BtnXMLload.Size = New System.Drawing.Size(140, 23)
        Me.BtnXMLload.TabIndex = 3
        Me.BtnXMLload.Text = "XML laden"
        Me.BtnXMLload.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnXMLload.UseVisualStyleBackColor = True
        '
        'BtnXMLsave
        '
        Me.BtnXMLsave.Image = Global.wpContact_Outlook_Add_in.My.Resources.Resources.SQLDatabase_16x
        Me.BtnXMLsave.Location = New System.Drawing.Point(158, 278)
        Me.BtnXMLsave.Name = "BtnXMLsave"
        Me.BtnXMLsave.Size = New System.Drawing.Size(140, 23)
        Me.BtnXMLsave.TabIndex = 4
        Me.BtnXMLsave.Text = "XML speichern"
        Me.BtnXMLsave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnXMLsave.UseVisualStyleBackColor = True
        '
        'BtnSave
        '
        Me.BtnSave.Location = New System.Drawing.Point(304, 307)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(140, 38)
        Me.BtnSave.TabIndex = 5
        Me.BtnSave.Text = "Speichern"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'lblXMLStatus
        '
        Me.lblXMLStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblXMLStatus.Location = New System.Drawing.Point(12, 307)
        Me.lblXMLStatus.Name = "lblXMLStatus"
        Me.lblXMLStatus.Size = New System.Drawing.Size(286, 38)
        Me.lblXMLStatus.TabIndex = 6
        Me.lblXMLStatus.Text = "leer"
        Me.lblXMLStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FormConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(459, 356)
        Me.Controls.Add(Me.lblXMLStatus)
        Me.Controls.Add(Me.BtnSave)
        Me.Controls.Add(Me.BtnXMLsave)
        Me.Controls.Add(Me.BtnXMLload)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnCancel)
        Me.Name = "FormConfig"
        Me.Text = "Form.Config"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BtnCancel As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents tbSQLPassword As Windows.Forms.TextBox
    Friend WithEvents tbSQLUserID As Windows.Forms.TextBox
    Friend WithEvents tbSQLInitialCatalog As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents tbSQLDataSource As Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As Windows.Forms.SaveFileDialog
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents BtnSQLtest As Windows.Forms.Button
    Friend WithEvents BtnOlNew As Windows.Forms.Button
    Friend WithEvents tbOlContactGroupPrefix As Windows.Forms.TextBox
    Friend WithEvents tbOlFolder As Windows.Forms.TextBox
    Friend WithEvents BtnXMLload As Windows.Forms.Button
    Friend WithEvents BtnXMLsave As Windows.Forms.Button
    Friend WithEvents BtnSave As Windows.Forms.Button
    Friend WithEvents lblXMLStatus As Windows.Forms.Label
    Friend WithEvents lblSQLStatus As Windows.Forms.Label
End Class

Imports System.Data
Imports System.Windows.Forms

Public Class FormGremien
    Private Sub FormGremien_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            MyLog("### GREMIUMFORM-LOAD start")
            Me.KeyPreview = vbTrue
            Me.Text = "Mitarbeiten in Gremien - wpContacts Session Outlook Plug-In"

            Dim agr As DataTable = GetAllGremien()
            With CbGremien
                .DisplayMember = "grname"
                .ValueMember = "grnr"
                .DataSource = agr
            End With

            With DataGridView1
                .AllowUserToAddRows = False
                .EditMode = False
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End With
            MyLog($"Set CbGremien.SelectedIndex = {My.Settings.LastGremium}")
            CbGremien.SelectedIndex = My.Settings.LastGremium
        Catch ex As Exception
            MyError(ex, "FormGremien_Load")
        End Try

        '### https://stackoverflow.com/questions/20070305/want-to-assign-datacolumn-with-checkbox-type-to-a-gridview
        'Dim dt As DataTable = New DataTable()
        'Dim dr As DataRow
        'Dim dgchk As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
        ''dt.Columns.Add(New DataColumn("Selected", GetType(DataGridBoolColumn)))
        'dt.Columns.Add(dgchk)
        'dt.Columns.Add("TextBoxCol")

        'dr = dt.NewRow()
        'dr(0) = False
        'dr(1) = "Hello"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr(0) = True
        'dr(1) = "World"
        'dt.Rows.Add(dr)

        'DataGridView1.DataSource = dt
        MyLog("### GREMIUMFORM-LOAD end")
    End Sub

    Private Sub CbGremien_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CbGremien.SelectedIndexChanged
        Try
            Dim comboBox As ComboBox = CType(sender, ComboBox)
            'MyLog($"{comboBox.SelectedItem} {comboBox.SelectedValue} {comboBox.SelectedIndex}")
            MyLog($"Set My.Settings.LastGremium = {comboBox.SelectedValue}")
            My.Settings.LastGremium = comboBox.SelectedValue
            My.Settings.Save()
            MyLog($"Looking for Members in {comboBox.SelectedValue}")
            Dim tbl As DataTable = ReadMembersOfGremium(comboBox.SelectedValue)
            DataGridView1.DataSource = tbl
            '### TODO Resize form with width of grid
            'DataGridView1.AutoResizeColumns()
        Catch ex As Exception
            MyError(ex, "CbGremien_SelectedIndexChanged")
        End Try
    End Sub

    Private Sub FormGremien_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class
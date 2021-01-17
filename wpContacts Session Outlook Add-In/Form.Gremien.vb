Imports System.Data
Imports System.Windows.Forms

Public Class FormGremien
    Private Sub FormGremien_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim agr As DataTable = GetAllGremien()
        With CbGremien
            .DisplayMember = "grname"
            .ValueMember = "grnr"
            .DataSource = agr
        End With
        PrintValues(agr, "AllGremien")

        'CbGremien.SelectedIndex = 1
        'CbGremien.Items.IndexOf(1)
        'comboBox.SelectedIndex = comboBox.Items.IndexOf("Director");

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
    End Sub

    Private Sub CbGremien_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CbGremien.SelectedIndexChanged
        Dim comboBox As ComboBox = CType(sender, ComboBox)
        Dim selectedEmployee = CType(comboBox.SelectedItem, String)
        Dim count As Integer = 0
        Dim resultIndex As Integer = -1
        resultIndex = comboBox.FindStringExact(comboBox.SelectedItem)
        While (resultIndex <> -1)
            comboBox.Items.RemoveAt(resultIndex)
            count += 1
            resultIndex = comboBox.FindStringExact(selectedEmployee, resultIndex)
        End While
        TextBox1.Text = TextBox1.Text & Microsoft.VisualBasic.vbCrLf & selectedEmployee & ": " & count
    End Sub
End Class
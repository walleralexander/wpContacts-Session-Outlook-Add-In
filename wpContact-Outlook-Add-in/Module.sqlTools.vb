Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Module SqlToolsModule
    Public Function SqlTestConnection(sqlds As String, sqlic As String, sqluid As String, sqlpwd As String) As Boolean
        Dim builder As New SqlConnectionStringBuilder With {
            .DataSource = sqlds,
            .InitialCatalog = sqlic,
            .UserID = sqluid,
            .Password = sqlpwd
        }
        Dim connectionstring As String = builder.ConnectionString
        Dim con As SqlConnection = Nothing
        If sqlds = "" Then
            MyLog("SqlTestConnection: SQL-credentials missing", "error")
            Return False
        End If
        Try
            Cursor.Current = Cursors.WaitCursor
            con = New SqlConnection(connectionstring)
            MyLog("SqlTestConnection Connection established")
            con.Open()
            Return True
        Catch ex As Exception
            MyLog("SqlTestConnection Error:" & ex.Message)
            Return False
        Finally
            If con.State <> ConnectionState.Closed Then con.Close()
            Cursor.Current = Cursors.WaitCursor
        End Try
    End Function

    Public Function SqlBuildConnection(sqlds As String, sqlic As String, sqluid As String, sqlpwd As String)
        Dim con As SqlConnection = Nothing
        Dim connectionstring As String
        Dim builder As New SqlConnectionStringBuilder With {
            .DataSource = sqlds,
            .InitialCatalog = sqlic,
            .UserID = sqluid,
            .Password = sqlpwd
        }
        connectionstring = builder.ConnectionString
        Try
            con = New SqlConnection(connectionstring)
            MyLog("SQL-Connection established")
            Return con
        Catch ex As Exception
            MyLog("SQL-Connection-Error:" & ex.Message)
        Finally
            If con.State <> ConnectionState.Closed Then con.Close()
        End Try
        Return Nothing
    End Function
    Function SqlFillDataTable_alt(ByRef con As SqlConnection, tablename As String, ByRef myDatatable As Object, Optional myOrder As String = "", Optional myWhere As String = "")
        Dim mySqlDataAdpter As SqlDataAdapter
        Dim myAnzRecords As Int16
        Dim sql As String = "select * from " & tablename
        If myOrder.Length > 0 Then sql &= " " & myOrder
        If myWhere.Length > 0 Then sql &= " " & myWhere
        Dim cleansql As String = sql.Replace("  ", "").Replace(vbCrLf, " ").Replace(vbTab, " ")
        MyLog($"Tabelle '{tablename}' sql: {cleansql}")
        mySqlDataAdpter = New SqlDataAdapter(sql, con)
        Try
            myAnzRecords = mySqlDataAdpter.Fill(myDatatable)
            MyLog($"SqlFillDataTable::Tabelle '{tablename}' gelesen: {myAnzRecords} Datensätze")
            myDatatable.TableName = tablename
        Catch ex As Exception
            MyLog($"Fehler: Tabelle '{tablename}' Message: {ex.Message}")
        End Try
        Return myAnzRecords
    End Function
    Function SqlFillDataTable(ByRef con As SqlConnection, tablename As String, sql As String)
        Dim cleansql As String = sql.Replace("  ", "").Replace(vbCrLf, " ").Replace(vbTab, " ")
        MyLog($"SqlFillDataTable2::info Tabelle '{tablename}' sql: {cleansql}")
        Dim myStep As Int16 = 0
        Dim myAnzRecords As Int16
        Dim myDatatable As DataTable = New DataTable()
        Dim mySqlDataAdpter As SqlDataAdapter = New SqlDataAdapter(sql, con)
        Try
            myDatatable.TableName = tablename
            myAnzRecords = mySqlDataAdpter.Fill(myDatatable)
            MyLog($"SqlFillDataTable2::info Tabelle '{tablename}' gelesen: {myAnzRecords} Datensätze")
        Catch ex As Exception
            MyLog($"SqlFillDataTable2:Fehler Tabelle '{tablename}' Step: {myStep} Message: {ex.Message}")
        End Try
        Return myDatatable
    End Function
End Module

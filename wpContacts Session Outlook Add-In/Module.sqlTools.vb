﻿Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms

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

Module SqlToolsModule
    Public debug As Boolean = My.Settings.Debug

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
            MyLog("SqlTestConnection: SQL-credentials missing")
            Return False
        End If
        Try
            Cursor.Current = Cursors.WaitCursor
            con = New SqlConnection(connectionstring)
            MyLog("SqlTestConnection Connection established")
            con.Open()
            Return True
        Catch ex As Exception
            MyError(ex, "SqlTestConnection")
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
            MyError(ex, "SqlBuildConnection")
        Finally
            If con.State <> ConnectionState.Closed Then con.Close()
        End Try
        Return Nothing
    End Function

    Function SqlFillDataTable(ByRef con As SqlConnection, tablename As String, sql As String)
        Dim cleansql As String = sql.Replace("  ", "").Replace(vbCrLf, " ").Replace(vbTab, " ")
        MyLog($"SqlFillDataTable::info Tabelle '{tablename}' sql: {cleansql}")
        Dim myAnzRecords As Int16
        Dim myDatatable As DataTable = New DataTable()
        Dim mySqlDataAdpter As SqlDataAdapter = New SqlDataAdapter(sql, con)
        Try
            myDatatable.TableName = tablename
            myAnzRecords = mySqlDataAdpter.Fill(myDatatable)
            MyLog($"SqlFillDataTable::info Tabelle '{tablename}' Anz: {myAnzRecords}")
        Catch ex As Exception
            MyError(ex, $"SqlFillDataTable '{tablename}'")
        End Try
        Return myDatatable
    End Function

    Function SqlFillDataTable2(sql As String, tablename As String)
        Dim myDatatable As DataTable = Nothing
        Dim myAnzRecords As Int16
        Try
            MyLog($"SqlFillDataTable2::info Tabelle '{tablename}' sql: {sql.Replace("  ", "").Replace(vbCrLf, " ").Replace(vbTab, " ")}")
            Dim con As SqlConnection = SqlBuildConnection(My.Settings.SQLDataSource, My.Settings.SQLInitialCatalog, My.Settings.SQLUserID, My.Settings.SQLPassword)
            Dim mySqlDataAdpter As SqlDataAdapter = New SqlDataAdapter(sql, con)
            myDatatable = New DataTable With {
                .TableName = tablename
            }
            myAnzRecords = mySqlDataAdpter.Fill(myDatatable) 'Hinzufügen der Daten und Anzahl als Rückgabewert
            MyLog($"SqlFillDataTable2::info Tabelle '{tablename}' gelesen: {myAnzRecords} Datensätze")
        Catch ex As Exception
            MyError(ex, "SqlFillDataTable2")
        End Try
        Return myDatatable
    End Function

End Module

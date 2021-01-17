Imports System.Data
Imports System.Data.SqlClient

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

Module SessionsModule
    ReadOnly SessionDateOffset As Int32 = 2415019
    Function Clean(mystring As Object)
        Return mystring.ToString.Trim
    End Function

    Function ReadContactsFromSession(ByRef con As SqlConnection)
        Dim tbl As DataTable = Nothing
        Try
            Dim Sql As String = "
                select tad.*, tpe.*, tsb.*, tat.*, tan.* from tad
                inner join tpe on tad.adnr = tpe.peadnr
                inner join tsb on tsb.sbnr = tpe.pesbnr
                inner join tat on tat.atnr = tsb.sbatnr
                inner join tan on tan.annr = tad.adannr
                where peedat = 0 and adnr > 0
                order by tad.adname
                "
            tbl = SqlFillDataTable(con, "SessionContacts", Sql)
            MyLog($"ReadContactsFromSession::OK:Tabelle 'tad' Contacts: {tbl.Rows.Count}")
        Catch ex As Exception
            MyLog($"ReadContactsFromSession::Fehler:Tabelle 'tad' Message: {ex.Message}")
        End Try
        Return tbl
    End Function
    Function ReadGremienFromSession(ByRef con As SqlConnection)
        Dim tbl As DataTable = Nothing
        Try
            Dim Sql As String = "
                select * from tgr
                inner join ttx on tgr.grtxnr=ttx.txnr
                where gredat = 0 and grnr > 0
                "
            tbl = SqlFillDataTable(con, "Gremien", Sql)
            MyLog($"ReadGremienFromSession::OK:Tabelle 'tad' Contacts: {tbl.Rows.Count}")
        Catch ex As Exception
            MyLog($"ReadGremienFromSession::Fehler:Tabelle 'tad' Message: {ex.Message}")
        End Try
        Return tbl
    End Function
    Function ReadOneGremiumFromSession(grnr As Int16)
        Dim tbl As DataTable = Nothing
        Dim con As SqlConnection = SqlBuildConnection(My.Settings.SQLDataSource, My.Settings.SQLInitialCatalog, My.Settings.SQLUserID, My.Settings.SQLPassword)
        Try
            Dim Sql As String = "
                select 
                mgnr, mggrnr, mgamnr, mgpenr, mgpunr, mgsort, mgfunk,
                grname, grkurz, amname, amstat, 
                peadnr, pepartei, pesch, pegdat, peedat,
                adname, advname, adtit, adtitn, adakgrad, adfirma1, adabt,  adberuf, adannr, adbvnr, 
                ademail, ademail2, adstr, adhnr, adplz, adort, adtel, adtel2, adtel3, adtel4, adtel5, adtel6
                from tmg
                inner join tgr on tmg.mggrnr=tgr.grnr
                inner join tam on tmg.mgamnr=tam.amnr
                inner join tpe on tmg.mgpenr=tpe.penr
                inner join tad on tpe.peadnr=tad.adnr
                where mggrnr = " & grnr.ToString & " and mgedat = 0 order by mgsort desc
            "
            tbl = SqlFillDataTable(con, "Gremien", Sql)
        Catch ex As Exception
            MyLog($"ReadOneGremiumFromSession::Fehler: Tabelle 'tad' Message: {ex.Message}")
        End Try
        Return tbl
    End Function
    Public Function GetMyGremien(adnr As String) As DataTable
        Dim tbl As DataTable = Nothing
        ' SessionDateOffset = 2415019
        ' mgadat = 2457122
        ' diff = 42103
        ' mgadat = 09.04.2015
        Try
            Dim Sql As String = "
                select grname, grkurz, mgfunk, mgadat, mgedat, pepartei
                    from tmg
                    join tpe on tmg.mgpenr=tpe.penr
                    join tgr on tmg.mggrnr=tgr.grnr
                    where peadnr=" & adnr.ToString & "
                    order by mgedat, mgadat, grname
            "
            'order by grname, mgadat
            tbl = SqlFillDataTable2(Sql, "MyGremien")
        Catch ex As Exception
            MyError(ex, "GetMyGremien")
        End Try
        Return tbl
    End Function
    Public Function GetMyKategorien(adnr As String)
        '# insert into ttx (txnr,txadat,txdef,txdef1,txdef2,txdef3,txdef4,txdef5,txdef6,txedat,txname,txsort,txtext,txtext1,txtyp,txwww,tx__nr,tx__typ,txcreated,txmodified) 
        '  values(32, 2459123, 0, 0, 0, 0, 0, 0, 0, 0,'Amt (Mitarbeiter/intern)',0,' ',' ',0,16,0,891,'2020-09-30 09:39:37.6600000','2020-09-30 12:07:18.6933333');
        '# insert into tgr 

        ' tgr ist nicht relevant

    End Function
    Public Function SaveConfigToXML(myfile)
        Try
            Dim configAPP As New DataSet("wpContactTool")
            Dim config As New DataTable("config")
            Dim configrow As DataRow
            config.Columns.Add(New DataColumn("ConfigName", GetType(System.String)))
            config.Columns.Add(New DataColumn("SQLDataSource", GetType(System.String)))
            config.Columns.Add(New DataColumn("SQLInitialCatalog", GetType(System.String)))
            config.Columns.Add(New DataColumn("SQLUserID", GetType(System.String)))
            config.Columns.Add(New DataColumn("SQLPassword", GetType(System.String)))
            config.Columns.Add(New DataColumn("olContactGroupPrefix", GetType(System.String)))
            config.Columns.Add(New DataColumn("olFolder", GetType(System.String)))
            'config.Columns.Add(New DataColumn("myLogfile", GetType(System.String)))
            config.PrimaryKey = New DataColumn() {config.Columns(0)}
            configrow = config.NewRow()
            For Each col As DataColumn In config.Columns
                If col.ColumnName = "ConfigName" Then
                    configrow(col.ColumnName) = "Config1"
                Else
                    configrow(col.ColumnName) = CallByName(My.Settings, col.ColumnName, CallType.Get, Nothing)
                End If
            Next
            config.Rows.Add(configrow)
            configAPP.Tables.Add(config)
            MyLog("write to" & myfile)
            configAPP.WriteXml(myfile)
        Catch ex As Exception
            MyLog(ex.Message, "error")
        End Try
        Return True
    End Function

    Function SaveContactsFromSession(dt As DataTable, file As String)
        Try
            MyLog($"SaveContactsFromSession::Info: Outputfile '{file}' ")
            dt.WriteXml(file)
        Catch ex As Exception
            MyLog($"SaveContactsFromSession::Fehler: SaveContactsFromSession '{file}' Message: {ex.Message}")
            Return False
        End Try
        Return True
    End Function
End Module

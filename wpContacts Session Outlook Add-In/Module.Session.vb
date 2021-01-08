Imports System.Data
Imports System.Data.SqlClient
Module SessionsModule
    Function Clean(mystring As Object)
        Return mystring.ToString.Trim
    End Function

    Function ReadContactsFromSession(ByRef con As SqlConnection)
        Dim tbl As DataTable = Nothing
        Try
            Dim Sql As String = "
                select tad.*, tpe.*, tsb.*, tat.* from tad
                inner join tpe on tad.adnr = tpe.peadnr
                inner join tsb on tsb.sbnr = tpe.pesbnr
                inner join tat on tat.atnr = tsb.sbatnr
                where peedat = 0 and adnr > 0
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

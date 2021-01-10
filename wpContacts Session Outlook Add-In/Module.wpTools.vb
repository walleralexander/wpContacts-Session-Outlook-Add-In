Imports System.Data
Imports log4net
'### https://www.codeproject.com/Tips/309804/Beginners-Guide-How-to-Implement-Log4net-in-VB-NET
'### https://andydunkel.net/2018/08/23/logging-mit-log4net/

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

Module WpToolsModule
    Public log As log4net.ILog = LogManager.GetLogger("WpTestFrameworkModule")
    Public Sub MyLog(message As String, Optional level As String = "info")

        Try
            If level <> "info" Then
                log.Error(message)
            Else
                log.Info(message)
            End If
        Catch ex As Exception
            log.Error(ex.Message)
        End Try
    End Sub
    Public Function ListSettings() As Boolean
        For Each Val As Configuration.SettingsProperty In My.Settings.Properties
            MyLog($"Name:{Val.Name} Type:'{Val.PropertyType}' Default:'{Val.DefaultValue}' Value:'{My.Settings.Item(Val.Name)}'")
        Next
        Return True
    End Function
    Public Function SetNetworkConfigPath()
        '### DefaultConfigFile in My.Settings an das Projekt anpassen
        'My.Settings.DefaultConfigFile = "wpContacts-Session-Outlook-Addin.xml"
        Try
            If Environment.GetEnvironmentVariable("wpconfig") > "" Then
                My.Settings.NetworkConfigStore = Environment.GetEnvironmentVariable("wpconfig")
            Else
                '### Thats for developement only
                'My.Settings.NetworkConfigStore = "\\v-046\msi\webpoint"
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function ReadConfigFromXML(myfile)
        Try
            MyLog("Reading XML config file")
            Dim configset As DataSet = New DataSet
            Dim configtable As DataTable
            Dim configrow As DataRow
            configset.ReadXml(myfile)
            configtable = configset.Tables("config")
            Dim config = Nothing
            For Each configrow In configtable.Rows
                Dim row As New Dictionary(Of String, String)
                For Each dc As DataColumn In configtable.Columns
                    'MyLog(dc.ColumnName & ": " & configrow(dc.ColumnName))
                    row.Add(dc.ColumnName, configrow(dc.ColumnName))
                Next
                config = row
            Next
            Return config
        Catch ex As Exception
            MyLog(ex.Message, "error")
            Return Nothing
        End Try
    End Function

    Private Sub PrintValues(ByVal table As DataTable, ByVal label As String)
        Console.WriteLine(label)
        For Each row As DataRow In table.Rows
            For Each column As DataColumn In table.Columns
                Console.Write("{0}{1}", ControlChars.Tab, row(column))
            Next column
            Console.WriteLine()
        Next row
    End Sub

End Module

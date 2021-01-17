Public Class ThisAddIn

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

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        MyLog("ADDINSTARTUP ##################################################################################")
        If Environment.GetEnvironmentVariable("wpdebug") > "" Then
            My.Settings.Debug = True
        Else
            My.Settings.Debug = False
        End If
        My.Settings.Save()
        MyLog($"DEBUG: {My.Settings.Debug} wpdebug={Environment.GetEnvironmentVariable("wpdebug")}")
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        MyLog("ADDINSHUTDOWN ################################################################################")
    End Sub
End Class

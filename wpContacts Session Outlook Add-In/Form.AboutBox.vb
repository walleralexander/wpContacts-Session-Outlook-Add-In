Public NotInheritable Class AboutBox1

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

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Me.Text = String.Format("Info {0}", ApplicationTitle)
        Me.LabelProductName.Text = My.Application.Info.ProductName
        Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = My.Application.Info.Description
        'Me.ProgPath.Text = My.Application.Info.DirectoryPath
        Me.ProgPath.Text = My.User.Name
        Me.EnvPath.Text = My.Settings.NetworkConfigStore
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

End Class

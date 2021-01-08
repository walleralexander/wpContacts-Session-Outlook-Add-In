Imports System.Deployment.Application
Imports System.Windows.Forms
Module DeploymentModule
    Public Function InstallUpdateSyncWithInfo()
        Dim info As UpdateCheckInfo
        If ApplicationDeployment.IsNetworkDeployed Then
            Dim ad As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
            Try
                info = ad.CheckForDetailedUpdate()
            Catch dde As DeploymentDownloadException
                MessageBox.Show("The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error: " + dde.Message)
                Return False
            Catch ide As InvalidDeploymentException
                MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message)
                Return False
            Catch ioe As InvalidOperationException
                MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message)
                Return False
            End Try

            If info.UpdateAvailable Then
                Dim doUpdate As Boolean = True
                If Not info.IsUpdateRequired Then
                    Dim dr As DialogResult = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available", MessageBoxButtons.OKCancel)
                    If Not DialogResult.OK = dr Then
                        doUpdate = False
                    End If
                Else
                    'Display a message that the app MUST reboot. Display the minimum required version.
                    MessageBox.Show("This application has detected a mandatory update from your current " &
                    "version to version " + info.MinimumRequiredVersion.ToString() & ". The application will now install the update and restart.",
                    "Update Available", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                If doUpdate Then
                    Try
                        'ad.Update()
                        MessageBox.Show("The application has been upgraded, and will now restart.")
                        'Application.Restart()
                    Catch dde As DeploymentDownloadException
                        MessageBox.Show("Cannot install the latest version of the application. \n\nPlease check your network connection, or try again later. Error: " & dde.Message)
                        Return False
                    End Try
                End If
            End If
        End If
        Return False
    End Function
End Module

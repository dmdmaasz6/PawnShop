Public Class frmConfig

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Dim bChanges As Boolean

        bChanges = False
        If txtCashRegister.Text <> p_CashRegister Then
            WriteRegistryCR(txtCashRegister.Text)
            bChanges = True
        End If

        If txtDBLocation.Text <> p_DBLocation Then
            WriteRegistryDB(txtDBLocation.Text)
            bChanges = True
        End If

        If txtAgreeLocation.Text <> "Not set using: " & p_AgreementPath And txtAgreeLocation.Text <> p_AgreementPath Then
            WriteRegistryAgreePath(txtAgreeLocation.Text)
            bChanges = True
        End If

        If txtShop.Text <> p_Branch Then
            WriteRegistryBrnch(txtShop.Text)
            bChanges = True
        End If

        If p_Audit = False Then
            If rbnAuditY.Checked = True Then
                WriteRegistryAudit("Y")
                bChanges = True
            End If
        Else
            If rbnAuditN.Checked = True Then
                WriteRegistryAudit("N")
                bChanges = True
            End If
        End If

        If bChanges = False Then
            MsgBox("No data was changed.", vbInformation)
        Else
            MsgBox("Application needs to be restarted, for any changes to take affect!", vbExclamation)
        End If

        Me.Close()

    End Sub

    Private Sub frmConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FolderBrowserDialog1.Description = "Select the path to save the Agreement Contracts." & vbCrLf & "Subfolders with the current date of agreement will automatically be created."
        If p_Audit = False Then
            rbnAuditN.Checked = True
            rbnAuditY.Checked = False
        Else
            rbnAuditN.Checked = False
            rbnAuditY.Checked = True
        End If

        txtCashRegister.Text = p_CashRegister
        txtDBLocation.Text = p_DBLocation

        If p_AgreementPath = "C:\Agreements" Then
            txtAgreeLocation.Text = "Not set using: " & p_AgreementPath
        Else
            txtAgreeLocation.Text = p_AgreementPath
        End If

        txtShop.Text = p_Branch

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub cmdFile_Click(sender As Object, e As EventArgs) Handles cmdFile.Click
        OpenFileDialog1.InitialDirectory = Application.StartupPath
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "All files (*.*)|*.*|mdb files (*.mdb)|*.mdb"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.RestoreDirectory = True

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            txtDBLocation.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub cmdFile2_Click(sender As Object, e As EventArgs) Handles cmdFile2.Click
        Me.FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop

        If FolderBrowserDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            txtAgreeLocation.Text = FolderBrowserDialog1.SelectedPath
        End If

    End Sub

End Class
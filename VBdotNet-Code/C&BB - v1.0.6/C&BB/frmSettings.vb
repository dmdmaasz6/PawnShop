Public Class frmSettings

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Me.Close()

    End Sub

    Private Sub frmSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtCashRegister.Text = p_CashRegister
        txtDBLocation.Text = p_DBLocation

        If p_AgreementPath = "C:\Agreements" Then
            txtAgreeLocation.Text = "Not set using: " & p_AgreementPath
        Else
            txtAgreeLocation.Text = p_AgreementPath
        End If

        txtShop.Text = p_Branch
        txtDate.Text = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
        txtCurrency.Text = p_Currency
        If p_Audit = False Then
            rbnAuditN.Checked = True
            rbnAuditY.Checked = False
        Else
            rbnAuditN.Checked = False
            rbnAuditY.Checked = True
        End If

    End Sub
End Class
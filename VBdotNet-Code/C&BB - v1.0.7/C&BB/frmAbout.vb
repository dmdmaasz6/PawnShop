Public Class frmAbout

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Me.Close()

    End Sub

    Private Sub frmAbout_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblName.Text = Replace(Application.ProductName, "&", "&&")
        lblVesion2.Text = Application.ProductVersion
        lblCompany.Text = Replace(Application.CompanyName, "&", "&&")

    End Sub
End Class
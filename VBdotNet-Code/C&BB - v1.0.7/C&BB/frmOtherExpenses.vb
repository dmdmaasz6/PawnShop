Public Class frmOtherExpenses

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub txtAmount_LostFocus(sender As Object, e As EventArgs) Handles txtAmount.LostFocus
        SetCurr(txtAmount)

    End Sub

    Private Sub frmOtherExpenses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtAmount.Text = ""
        cboCategory.SelectedIndex = -1
        txtDescription.Text = ""
        'dtDate = Today

    End Sub

    Private Sub txtAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAmount.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Price must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If

    End Sub

    Private Sub cmdProcess_Click(sender As Object, e As EventArgs) Handles cmdProcess.Click
        ' Input validation
        If cboCategory.Text = "" Then
            MsgBox("Please specify a category.", vbInformation, "Missing Information")
            Exit Sub
        ElseIf Trim(txtAmount.Text) = "" Then
            MsgBox("Please specify an amount.", vbInformation, "Missing Information")
            Exit Sub
        End If

        ' Prepare Values
        Dim strDescription As String
        Dim intIndex As Integer
        If Trim(txtDescription.Text) = "" Then 'Error generated in insert statement when txtDescription is empty
            strDescription = "No description"
        Else
            strDescription = txtDescription.Text
        End If
        intIndex = cboCategory.SelectedIndex + 1 'ListIndex starts at 0

        Try
            'INSERT INTO CashTransactions (CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES (CashRegister,Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID)
            Me.CashTransactionsTableAdapter1.Insert("Categorized Expenses - index at top of report", FormatDateTime(dtDate.Value, DateFormat.ShortDate), 0, 0, 0, 0, intIndex, strDescription, -(FormatCurrency(txtAmount.Text)), 0, 0, 0, p_CashRegister)

            If p_Audit = True Then AuditToFile("Other Expenses:", "Date: " & FormatDateTime(dtDate.Value, DateFormat.ShortDate) & ", " & cboCategory.SelectedItem & ", " & strDescription & ", Amount: -" & (FormatCurrency(txtAmount.Text)))

            Me.Close()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub
End Class
Imports System.Data.OleDb

Public Class frmRptExpiredStock

    Private Sub frmRptExpiredStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadList()

    End Sub


    Private Sub ClearForm()
        Try
            dgExpiredPawnStock.ClearSelection()
            If dgExpiredPawnStock.Rows.Count <> 0 Then
                Do While dgExpiredPawnStock.Rows.Count <> 0
                    dgExpiredPawnStock.Rows.Remove(dgExpiredPawnStock.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try


    End Sub

    Private Sub LoadList()
        Dim table As New DataTable

        Try
            dgExpiredPawnStock.ClearSelection()

            'SELECT PawnStock.Quantity, PawnStock.Description, PawnStock.SerialNr, PawnStock.Price, PawnStock.PnCID, PawnStock.PSID FROM PawningTransactions INNER JOIN PawnStock ON PawningTransactions.PnCID = PawnStock.PnCID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <= ?) AND (PawningTransactions.Status = 1)
            com.CommandText = "SELECT PawnStock.PnCID, PawnStock.Quantity, PawnStock.Description, PawnStock.SerialNr, PawnStock.Price, PawnStock.PSID FROM PawningTransactions INNER JOIN PawnStock ON PawningTransactions.PnCID = PawnStock.PnCID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <=  #" & FormatDateTime(Now, DateFormat.ShortDate) & "#) AND (PawningTransactions.Status = 1)"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            table.Load(reader)

            With dgExpiredPawnStock
                .DataSource = table
            End With

            reader.Close()


            dgExpiredPawnStock.Columns("PnCID").HeaderText = "CBBID"
            dgExpiredPawnStock.Columns("Quantity").HeaderText = "Qty"
            dgExpiredPawnStock.Columns("Quantity").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgExpiredPawnStock.Columns("PSID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgExpiredPawnStock.Columns("Price").DefaultCellStyle.Format = p_Currency & "#,##0.00"
            dgExpiredPawnStock.Columns("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgExpiredPawnStock.AutoResizeColumnHeadersHeight(DataGridViewAutoSizeColumnMode.ColumnHeader)
            dgExpiredPawnStock.AutoResizeRowHeadersWidth(DataGridViewAutoSizeColumnMode.ColumnHeader)
            dgExpiredPawnStock.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

            dgExpiredPawnStock.Columns("Quantity").Width = 50
            dgExpiredPawnStock.Columns("Description").Width = 300
            dgExpiredPawnStock.Columns("SerialNr").Width = 150
            dgExpiredPawnStock.Columns("Price").Width = 100

            dgExpiredPawnStock.Refresh()

            reader.Close()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        ClearForm()
        Me.Close()

    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        With frmReports
            .Label1.Text = "Generating Expired Pawn Report ..."
            .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
            .Height = 67
            .Report = frmReports.ReportType.ExpiredPawn
            .ShowDialog(Me)
        End With
    End Sub

    Private Sub dgExpiredPawnStock_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgExpiredPawnStock.CellContentClick

    End Sub
End Class
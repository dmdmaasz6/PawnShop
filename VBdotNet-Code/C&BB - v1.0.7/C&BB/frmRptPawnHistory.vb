Imports System.Data.OleDb

Public Class frmRptPawnHistory
    Public sDateFrom As String
    Public sDateTo As String

    Private Sub frmRptPawnHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SELECT PawningTransactions.DateHandedIn, PawningTransactions.PnCID, Customers.Name, Customers.IDNr, PawningTransactions.PurchaseAmount, PawningTransactions.BuyBackAmount, PawningTransactions.ExpiryDate, PawningTransactions.BuyBackDate, PawningTransactions.Status FROM PawningTransactions INNER JOIN Customers ON (PawningTransactions.CustID = Customers.CustID) where (PawningTransactions.DateHandedIn between #" & sDateFrom & "# and #" & sDateTo & "#) order by PawningTransactions.DateHandedIn
        LoadList()
    End Sub

    Private Sub LoadList()
        Dim table As New DataTable
        Dim str(9) As String
        Dim itm As ListViewItem

        Try
            ListView1.Columns.Add("DateHandedIn")
            ListView1.Columns.Add("PnCID")
            ListView1.Columns.Add("Name")
            ListView1.Columns.Add("IDNr")
            ListView1.Columns.Add("PurchaseAmount")
            ListView1.Columns.Add("BuyBackAmount")
            ListView1.Columns.Add("ExpiryDate")
            ListView1.Columns.Add("BuyBackDate")
            ListView1.Columns.Add("Status")

            dgPawnHistory.ClearSelection()

            'SELECT PawnStock.Quantity, PawnStock.Description, PawnStock.SerialNr, PawnStock.Price, PawnStock.PnCID, PawnStock.PSID FROM PawningTransactions INNER JOIN PawnStock ON PawningTransactions.PnCID = PawnStock.PnCID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <= ?) AND (PawningTransactions.Status = 1)
            ' com.CommandText = "SELECT PawnStock.PnCID, PawnStock.Quantity, PawnStock.Description, PawnStock.SerialNr, PawnStock.Price, PawnStock.PSID FROM PawningTransactions INNER JOIN PawnStock ON PawningTransactions.PnCID = PawnStock.PnCID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <=  #" & FormatDateTime(Now, DateFormat.ShortDate) & "#) AND (PawningTransactions.Status = 1)"
            com.CommandText = "SELECT PawningTransactions.DateHandedIn, PawningTransactions.PnCID, Customers.Name, Customers.IDNr, PawningTransactions.PurchaseAmount, PawningTransactions.BuyBackAmount, PawningTransactions.ExpiryDate, PawningTransactions.BuyBackDate, PawningTransactions.Status FROM PawningTransactions INNER JOIN Customers ON (PawningTransactions.CustID = Customers.CustID) where (PawningTransactions.DateHandedIn between #" & sDateFrom & "# and #" & sDateTo & "#) order by PawningTransactions.DateHandedIn"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                str(0) = reader(0).ToString()
                str(1) = reader(1).ToString()
                str(2) = reader(2).ToString()
                str(3) = reader(3).ToString()
                str(4) = reader(4).ToString()
                str(5) = reader(5).ToString()
                str(6) = reader(6).ToString()
                str(7) = reader(7).ToString()
                str(8) = reader(8).ToString()
                itm = New ListViewItem(str)
                ListView1.Items.Add(itm)
            End While

            table.Load(reader)

            With dgPawnHistory
                .DataSource = table
            End With

            reader.Close()

            'dgPawnHistory.Columns("DateHandedIn").HeaderText = "DateHandedIn"
            'dgPawnHistory.Columns("Quantity").HeaderText = "Qty"
            'dgPawnHistory.Columns("Quantity").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'dgPawnHistory.Columns("PSID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            'dgPawnHistory.Columns("Price").DefaultCellStyle.Format = p_Currency & "#,##0.00"
            'dgPawnHistory.Columns("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            'dgPawnHistory.AutoResizeColumnHeadersHeight(DataGridViewAutoSizeColumnMode.ColumnHeader)
            'dgPawnHistory.AutoResizeRowHeadersWidth(DataGridViewAutoSizeColumnMode.ColumnHeader)
            'dgPawnHistory.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

            'dgPawnHistory.Columns("Quantity").Width = 50
            'dgPawnHistory.Columns("Description").Width = 300
            'dgPawnHistory.Columns("SerialNr").Width = 150
            'dgPawnHistory.Columns("Price").Width = 100

            dgPawnHistory.Refresh()
            ListView1.Refresh()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub ReportViewer1_Load(sender As Object, e As EventArgs)

    End Sub
End Class
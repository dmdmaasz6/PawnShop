Imports System.Data.OleDb

Public Class frmCustomerHistory

    Private Sub grpCustomer_Enter(sender As Object, e As EventArgs) Handles grpCustomer.Enter

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Try
            If dgCustHistory.Rows.Count <> 0 Then
                Do While dgCustHistory.Rows.Count <> 0
                    dgCustHistory.Rows.Remove(dgCustHistory.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

        Me.Close()

    End Sub

    Private Sub frmCustomerHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CBB_DataSet.PawnStock' table. You can move, or remove it, as needed.
        'Me.PawningTransactionsTableAdapter.CustHistory(Me.CBB_DataSet.PawningTransactions, CType(txtCustID.Text, Integer))
        SetupDataGridView()
        'DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgCustHistory.AutoResizeColumnHeadersHeight(DataGridViewAutoSizeColumnMode.ColumnHeader)
        dgCustHistory.AutoResizeRowHeadersWidth(DataGridViewAutoSizeColumnMode.ColumnHeader)
        dgCustHistory.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
        dgCustHistory.Refresh()


    End Sub

    Private Sub SetupDataGridView()

        dgCustHistory.ColumnCount = 13

        With dgCustHistory
            .Name = "Customer History"
            .AutoSizeRowsMode = _
                DataGridViewAutoSizeRowsMode.DisplayedCells

            .Columns(0).Name = "PnCID"
            .Columns(1).Name = "Description"
            .Columns(2).Name = "Serial Nr"
            .Columns(3).Name = "Month"
            .Columns(4).Name = "Time"
            .Columns(5).Name = "Amount"
            .Columns(6).Name = "DateHandedIn"
            .Columns(7).Name = "BuyBackAmount"
            .Columns(8).Name = "ExpiryDate"
            .Columns(9).Name = "Status"
            .Columns(10).Name = "ActualBuyBackAmount"
            .Columns(11).Name = "BuyBackDate"
            .Columns(12).Name = "ExpiryDate (7WorkDays)"

            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End With

        Try
            com.CommandText = "SELECT a.PnCID, b.Description, b.SerialNr, a.MonthNr, a.PTTime, a.PurchaseAmount, " _
                            & "a.DateHandedIn, a.BuyBackAmount, a.ExpiryDate, a.Status, a.ActualBuyBackAmount, a.BuyBackDate, " _
                            & "a.ExpiryDateAfter7WorkDays " _
                            & "FROM PawningTransactions a " _
                            & "INNER JOIN PawnStock b ON a.PnCID = b.PnCID " _
                            & "WHERE a.CustID = " & CType(txtCustID.Text, Integer)
            
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Me.dgCustHistory.Rows.Add(reader(0).ToString(), reader(1).ToString(), reader(2).ToString(), reader(3).ToString(), _
                                          Microsoft.VisualBasic.Right(reader(4).ToString(), 8), reader(5).ToString(), _
                                          Microsoft.VisualBasic.Left(reader(6).ToString(), 10), reader(7).ToString(), _
                                          Microsoft.VisualBasic.Left(reader(8).ToString(), 10), reader(9).ToString(), _
                                          reader(10).ToString(), Microsoft.VisualBasic.Left(reader(11).ToString(), 10), _
                                          Microsoft.VisualBasic.Left(reader(12).ToString(), 10))
            End While
            reader.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub


    Private Sub dgCustHistory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgCustHistory.CellContentClick

    End Sub
End Class
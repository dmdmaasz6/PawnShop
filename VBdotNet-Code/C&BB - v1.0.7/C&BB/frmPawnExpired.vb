Imports System.Data.OleDb

Public Class frmPawnExpired

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub ClearFields()


        cmdExit.Visible = False
        cmdReportAndMove.Visible = False
        cmdUpdate.Visible = False
        cmdExit2.Visible = True

        StatusStrip1.Width = Me.Width
        SS1.Text = ""

        If dgExpiredPawningStock.Rows.Count > 0 Then
            Do While dgExpiredPawningStock.Rows.Count <> 0
                dgExpiredPawningStock.Rows.Remove(dgExpiredPawningStock.Rows(0))
            Loop
        End If

        If dgExpiredPawningTransactions.Rows.Count > 0 Then
            Do While dgExpiredPawningTransactions.Rows.Count <> 0
                dgExpiredPawningTransactions.Rows.Remove(dgExpiredPawningTransactions.Rows(0))
            Loop
        End If


    End Sub

    Private Sub frmPawnExpired_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim table As New DataTable

        ClearFields()

        Try
            com.CommandText = "SELECT Customers.Name, PawningTransactions.DateHandedIn, PawningTransactions.ExpiryDateAfter7WorkDays, PawningTransactions.PurchaseAmount, PawningTransactions.BuybackAmount, PawningTransactions.PnCID, PawningTransactions.Status FROM Customers INNER JOIN PawningTransactions ON Customers.CustID = PawningTransactions.CustID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <= #" & FormatDateTime(Now, DateFormat.ShortDate) & "#) AND (PawningTransactions.Status = 1) ORDER BY Customers.Name, PawningTransactions.DateHandedIn"
            Dim reader As OleDbDataReader = com.ExecuteReader()
            table.Load(reader)

            With dgExpiredPawningTransactions
                .DataSource = table
            End With

            reader.Close()

            If dgExpiredPawningTransactions.RowCount = 0 Then
                MsgBox("No pawning transactions have expired, since you last checked", vbInformation, "No expired transactions")
                cmdExit.Visible = False
                cmdReportAndMove.Visible = False
                cmdUpdate.Visible = False
                cmdExit_Click(Nothing, Nothing)
                Exit Sub
            End If


            dgExpiredPawningTransactions.Columns("PnCID").Visible = False
            dgExpiredPawningTransactions.Columns("Status").Visible = False

            dgExpiredPawningTransactions.Columns("DateHandedIn").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgExpiredPawningTransactions.Columns("ExpiryDateAfter7WorkDays").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgExpiredPawningTransactions.Columns("PurchaseAmount").DefaultCellStyle.Format = p_Currency & "#,##0.00"
            dgExpiredPawningTransactions.Columns("PurchaseAmount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgExpiredPawningTransactions.Columns("BuybackAmount").DefaultCellStyle.Format = p_Currency & "#,##0.00"
            dgExpiredPawningTransactions.Columns("BuybackAmount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgExpiredPawningTransactions.AutoResizeColumnHeadersHeight(DataGridViewAutoSizeColumnMode.ColumnHeader)
            dgExpiredPawningTransactions.AutoResizeRowHeadersWidth(DataGridViewAutoSizeColumnMode.ColumnHeader)
            dgExpiredPawningTransactions.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

            dgExpiredPawningTransactions.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            dgExpiredPawningTransactions.ClearSelection()

            dgExpiredPawningTransactions.Refresh()


        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub dgExpiredPawningTransactions_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgExpiredPawningTransactions.CellContentClick
        Dim value As Object = dgExpiredPawningTransactions.Rows(e.RowIndex).Cells("PnCID").Value

        If dgExpiredPawningStock.Rows.Count > 0 Then
            Do While dgExpiredPawningStock.Rows.Count <> 0
                dgExpiredPawningStock.Rows.Remove(dgExpiredPawningStock.Rows(0))
            Loop
        End If

        If IsDBNull(value) Then
            txtPnCID.Text = ""
            SS1.Text = ""
            cmdExit.Visible = False
            cmdReportAndMove.Visible = False
            cmdUpdate.Visible = False
            cmdExit2.Visible = True
        Else
            SS1.Text = ""
            cmdExit.Visible = True
            cmdReportAndMove.Visible = True
            cmdUpdate.Visible = True
            cmdExit2.Visible = False
            txtPnCID.Text = CType(value, String)
        End If

    End Sub

    Private Sub txtPnCID_TextChanged(sender As Object, e As EventArgs) Handles txtPnCID.TextChanged
        Dim table As New DataTable
        Dim lPncId As Long

        Try
            If txtPnCID.Text <> "" Then
                lPncId = CLng(txtPnCID.Text)
                'Me.CustomersTableAdapter.FillBySearchCustByNameID(Me.CBB_DataSet.Customers, "%" & txtName.Text & "%", "%" & txtIDnr.Text & "%")
                'Me.PawnStockTableAdapter.ExpiredPawnStock(Me.CBB_DataSet.PawnStock, lPncId)

                com.CommandText = "SELECT Quantity, Description, SerialNr, Price, PnCID, Status FROM PawnStock where PnCID =" & lPncId
                Dim reader As OleDbDataReader = com.ExecuteReader()
                table.Load(reader)

                With dgExpiredPawningStock
                    .DataSource = table
                End With

                reader.Close()

                dgExpiredPawningStock.Columns("PnCID").Visible = False
                dgExpiredPawningStock.Columns("Status").Visible = False

                dgExpiredPawningStock.Columns("Quantity").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgExpiredPawningStock.Columns("Price").DefaultCellStyle.Format = p_Currency & "#,##0.00"
                dgExpiredPawningStock.Columns("Price").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                dgExpiredPawningStock.AutoResizeColumnHeadersHeight(DataGridViewAutoSizeColumnMode.ColumnHeader)
                dgExpiredPawningStock.AutoResizeRowHeadersWidth(DataGridViewAutoSizeColumnMode.ColumnHeader)
                dgExpiredPawningStock.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

                dgExpiredPawningStock.Columns("Quantity").Width = 50
                dgExpiredPawningStock.Columns("Description").Width = 300
                dgExpiredPawningStock.Columns("SerialNr").Width = 150
                dgExpiredPawningStock.Columns("Price").Width = 100

                dgExpiredPawningStock.SelectionMode = DataGridViewSelectionMode.FullRowSelect

                dgExpiredPawningStock.ClearSelection()
                dgExpiredPawningStock.ReadOnly = False

                dgExpiredPawningStock.Refresh()

            End If

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try



    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        Dim value_Quantity As Object
        Dim value_Description As Object
        Dim value_SerialNr As Object
        Dim value_Price As Object
        Dim value_PnCID As Object
        Dim iRowCount As Integer
        Dim i As Integer

        Try


            If dgExpiredPawningStock.RowCount > 0 Then
                iRowCount = dgExpiredPawningStock.RowCount
                For i = 0 To iRowCount - 1

                    value_Quantity = dgExpiredPawningStock.Rows(i).Cells("Quantity").Value
                    value_Description = dgExpiredPawningStock.Rows(i).Cells("Description").Value
                    value_SerialNr = dgExpiredPawningStock.Rows(i).Cells("SerialNr").Value
                    value_Price = dgExpiredPawningStock.Rows(i).Cells("Price").Value
                    value_PnCID = dgExpiredPawningStock.Rows(i).Cells("PnCID").Value

                    'MsgBox(value_PnCID & vbCrLf & value_Description & vbCrLf & value_SerialNr & vbCrLf & value_Price)

                    com.CommandText = "UPDATE PawnStock SET Quantity = " & value_Quantity & ", Description = '" & value_Description & "', SerialNr = '" & value_SerialNr & "', Price = " & value_Price & " WHERE PnCID = " & value_PnCID
                    'com.CommandText = "SELECT Quantity, Description, SerialNr, Price, PnCID, Status FROM PawnStock where PnCID =" & lPncId
                    Dim reader As OleDbDataReader = com.ExecuteReader()
                    reader.Close()

                    SS1.Text = value_Description & " updated ..."

                Next i
            End If

            txtPnCID_TextChanged(Nothing, Nothing)

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cmdReportAndMove_Click(sender As Object, e As EventArgs) Handles cmdReportAndMove.Click
        Dim i As Integer

        Try
            'Generate Report:
            'SELECT PawnStock.Quantity, PawnStock.Description, PawnStock.SerialNr, PawnStock.Price, PawnStock.PnCID, PawnStock.PSID FROM PawningTransactions INNER JOIN PawnStock ON PawningTransactions.PnCID = PawnStock.PnCID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <= ?) AND (PawningTransactions.Status = 1)
            Me.Cursor = Cursors.WaitCursor
            frmRptExpiredStock.ShowDialog(Me)
            Me.Cursor = Cursors.Default

            '**'Set Status Values of PawningTransactions to 3 (=Expired)
            Dim reader As OleDbDataReader
            For i = 0 To dgExpiredPawningTransactions.RowCount - 1
                com.CommandText = "UPDATE PawningTransactions SET Status = '3' where PnCID = " & dgExpiredPawningTransactions.Rows(i).Cells("PnCID").Value   ' (SELECT PnCID FROM PawningTransactions WHERE ExpiryDateAfter7WorkDays <= #" & FormatDateTime(Now, DateFormat.ShortDate) & "#) AND Status = 1)"
                reader = com.ExecuteReader()
                reader.Close()
            Next i

            frmPawnExpired_Load(Nothing, Nothing)

        Catch ex As Exception
            MsgBox("An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
                    "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
                    "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
                    "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
                    "4. If you get this same error, inform the programmer." & vbCrLf & vbCrLf & _
                    "Error Detail:" & vbCrLf & _
                    ex.Message, vbCritical, "Unsuccessful database update.")
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)

        End Try

    End Sub
End Class
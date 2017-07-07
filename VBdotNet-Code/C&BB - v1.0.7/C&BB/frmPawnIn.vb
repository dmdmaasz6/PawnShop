Imports System.Data.OleDb

Public Class frmPawnIn

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub ClearFields()
        ' Clear populated fields

        ClearFields2()

        txtName.Text = ""
        txtIDnr.Text = ""

    End Sub

    Private Sub ClearFields2()
        ' Clear populated fields

        ClearDetail()
        txtCustID.Text = ""
        txtPNCID.Text = ""
        dgPawnInList.Visible = False
        lblNoTransactions.Visible = False

        Try
            If dgCustomerList.Rows.Count > 0 Then
                Do While dgCustomerList.Rows.Count <> 0
                    dgCustomerList.Rows.Remove(dgCustomerList.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

        txtPNCID.Focus()

    End Sub

    Private Sub ClearDetail()
        ' Clear populated fields
        lblCustomerInformation.Text = ""
        lblGoods.Text = ""
        lblPawningTransaction.Text = ""
        txtActualBuyBackAmount.Text = ""

    End Sub


    Private Sub frmPawnIn_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ClearFields()

    End Sub

    Private Sub cmdDetail_Click(sender As Object, e As EventArgs) Handles cmdDetail.Click
        Dim sSQL As String
        Dim bReader As Boolean
        Dim bDisplayMsg As Boolean

        bDisplayMsg = False

        ClearDetail()

        ' Check if PNC has been given
        If Trim(txtPNCID.Text) = "" Then
            MsgBox("Must specify a CBB-ID.", vbCritical, "Missing information")
            ClearFields()
            Exit Sub
        End If

        bReader = False
        '0 PawningTransactions.PnCID, 
        '1 PawningTransactions.MonthNr, 
        '2 PawningTransactions.PTTime,
        '3 Customers.CustID, 
        '4 Customers.Name, 
        '5 Customers.TelW, 
        '6 Customers.IDNr, 
        '7 Customers.TelH, 
        '8 Customers.Address, 
        '9 PawnStock.PSID, 
        '10 PawnStock.Quantity, 
        '11 PawnStock.Description, 
        '12 PawnStock.SerialNr, 
        '13 PawningTransactions.PurchaseAmount, 
        '14 PawningTransactions.DateHandedIn,
        '15 PawningTransactions.BuyBackAmount, 
        '16 PawningTransactions.Status, 
        '17 PawningTransactions.ExpiryDate, 
        '18 PawningTransactions.ExpiryDateAfter7WorkDays
        sSQL = "SELECT PawningTransactions.PnCID, PawningTransactions.MonthNr, PawningTransactions.PTTime, Customers.CustID, Customers.Name, Customers.TelW, Customers.IDNr, Customers.TelH, Customers.Address, PawnStock.PSID, PawnStock.Quantity, PawnStock.Description, PawnStock.SerialNr, PawningTransactions.PurchaseAmount, PawningTransactions.DateHandedIn, PawningTransactions.BuyBackAmount, PawningTransactions.Status, PawningTransactions.ExpiryDate, PawningTransactions.ExpiryDateAfter7WorkDays FROM (Customers INNER JOIN PawningTransactions ON Customers.CustID = PawningTransactions.CustID) INNER JOIN PawnStock ON PawningTransactions.PnCID = PawnStock.PnCID WHERE PawningTransactions.PnCID = " & txtPNCID.Text
        com.CommandText = sSQL
        Dim reader As OleDbDataReader = com.ExecuteReader()
        While reader.Read()
            bReader = True
            If reader(16).ToString = "2" Then 'Status
                MsgBox("This pawning transaction has already been bought back.", vbInformation, "Cannot proceed")
                txtPNCID.Text = ""
                reader.Close()
                Exit Sub
            End If

            ' Populate form controls
            lblCustomerInformation.Text = _
                "Name: " & reader(4).ToString & vbCrLf & _
                "IDNr: " & reader(6).ToString & vbCrLf & _
                "TelW: " & reader(5).ToString & vbCrLf & _
                "TelH: " & reader(7).ToString & vbCrLf & _
                "Address : " & reader(8).ToString

            lblPawningTransaction.Text = _
                "Date Handed In: " & FormatDateTime(reader(14).ToString, DateFormat.ShortDate) & vbCrLf & _
                "Purchase Amount: " & Format(reader(13).ToString, "currency") & vbCrLf & _
                "Expiry Date: " & FormatDateTime(reader(17).ToString, DateFormat.ShortDate) & vbCrLf & _
                "Buy Back Amount: " & Format(reader(15).ToString, "currency")

            txtActualBuyBackAmount.Text = reader(15).ToString

            ' If grace period has expired, notify
            If bDisplayMsg = False Then ' Do not display if message already displayed with first item.
                If CDate(FormatDateTime(reader(18).ToString, DateFormat.ShortDate)) < CDate(FormatDateTime(Now, DateFormat.ShortDate)) Then
                    MsgBox("The transaction is already past the 7 days grace period.", vbCritical, "Transaction Expired")
                    bDisplayMsg = True
                End If
            End If

            lblGoods.Text = lblGoods.Text & _
                reader(10).ToString & "x " & _
                reader(11).ToString & " (" & _
                reader(12).ToString & ")" & vbCrLf

        End While
        reader.Close()

        If bReader = False Then
            MsgBox("The CBB-ID is not valid", vbInformation, "Invalid CBB-ID")
            ClearFields()
        End If

    End Sub

    Private Sub cmdPawnIn_Click(sender As Object, e As EventArgs) Handles cmdPawnIn.Click
        Dim sSQL As String

        'Ensure that minimum information has been given
        If Trim(txtPNCID.Text) = "" Then
            MissingInfo("a CBB-ID")
            txtPNCID.Focus()
            Exit Sub
        ElseIf Trim(txtActualBuyBackAmount.Text) = "" Then
            MissingInfo("an actual buyback amount")
            txtActualBuyBackAmount.Focus()
            Exit Sub
        End If

        txtActualBuyBackAmount.Text = Replace(txtActualBuyBackAmount.Text, ",", "")

        'Ensure that given txtpncid is displayed
        If Trim(lblCustomerInformation.Text) = "" Then
            cmdDetail_Click(Nothing, Nothing)

            MsgBox("First review the information, to ensure correctness.", vbInformation, "Review")
            Exit Sub
        End If


        Try
            ' Change Status field in PawningTransactions Table
            sSQL = "UPDATE PawningTransactions SET PawningTransactions.Status = 2, PawningTransactions.ActualBuyBackAmount = " & txtActualBuyBackAmount.Text & ", PawningTransactions.BuyBackDate = #" & dtpBuyBackDate.Value & "# WHERE [PawningTransactions].[PnCID]= " & txtPNCID.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            reader.Close()
            If p_Audit = True Then AuditToFile("UPDATE PawningTransactions", "UPDATE PawningTransactions SET PawningTransactions.Status = 2, PawningTransactions.ActualBuyBackAmount = " & txtActualBuyBackAmount.Text & ", PawningTransactions.BuyBackDate = #" & dtpBuyBackDate.Value & "# WHERE [PawningTransactions].[PnCID]= " & txtPNCID.Text)

            ' Change Status field in PawnStock Table
            sSQL = "UPDATE PawnStock SET PawnStock.Status = 2 WHERE [PawnStock].[PnCID]=" & txtPNCID.Text
            com.CommandText = sSQL
            reader = com.ExecuteReader()
            reader.Close()
            If p_Audit = True Then AuditToFile("UPDATE PawnStock", "UPDATE PawnStock SET PawnStock.Status = 2 WHERE [PawnStock].[PnCID]=" & txtPNCID.Text)

            ' Insert into CashTransactions
            'dePNC.cmdInsertCashTransaction(gstrCR, "Pawning In", dtpBuyBackDate.Value, txtPNCID.Text, 0, 0, 0, 0, "", txtActualBuyBackAmount.Text, 0, 0, 0)
            'INSERT INTO CashTransactions (CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES (CashRegister,Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID)
            sSQL = "INSERT INTO CashTransactions (CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES " & _
                "('" & p_CashRegister & "', 'Pawning In', '" & FormatDateTime(dtpBuyBackDate.Value, DateFormat.ShortDate) & "', " & txtPNCID.Text & ", 0, 0, 0, 0, '', " & txtActualBuyBackAmount.Text & ", 0, 0, 0)"
            com.CommandText = sSQL
            reader = com.ExecuteReader()
            reader.Close()
            If p_Audit = True Then AuditToFile("Pawning In", "INSERT INTO CashTransactions " & FormatDateTime(dtpBuyBackDate.Value, DateFormat.ShortDate) & ", " & txtPNCID.Text & ", 0, 0, 0, 0, '', " & txtActualBuyBackAmount.Text & ", 0, 0, 0")

            MsgBox("Remember to have the customer sign the original CBB document", vbCritical, "IMPORTANT")

            Me.Close()



        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            MsgBox("An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
                    "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
                    "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
                    "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
                    "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update.")

        End Try

    End Sub

    Private Sub txtPNCID_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPNCID.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("ID must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If
    End Sub

    Private Sub txtActualBuyBackAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtActualBuyBackAmount.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("BuyBack Amount must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If
    End Sub

    Private Sub cmdSearchForPNC_Click(sender As Object, e As EventArgs) Handles cmdSearchForPNC.Click
        ClearFields2()

        Try
            If txtName.Text <> "" Or txtIDnr.Text <> "" Then
                Me.CustomersTableAdapter.FillBySearchCustByNameID(Me.CBB_DataSet.Customers, "%" & txtName.Text & "%", "%" & txtIDnr.Text & "%")
            Else
                MsgBox("Enter Name or ID Nr.", vbInformation + vbOKOnly, "No Vlaue Supplied!")
                Exit Sub
            End If
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try
        'SearchCustomer()

        dgCustomerList.AutoResizeColumnHeadersHeight(DataGridViewAutoSizeColumnMode.ColumnHeader)
        dgCustomerList.AutoResizeRowHeadersWidth(DataGridViewAutoSizeColumnMode.ColumnHeader)
        dgCustomerList.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
        dgCustomerList.Refresh()
        dgCustomerList.ClearSelection()
    End Sub

    Private Sub dgCustomerList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgCustomerList.CellContentClick
        Dim value As Object = dgCustomerList.Rows(e.RowIndex).Cells(0).Value

        If IsDBNull(value) Then
            txtCustID.Text = ""
        Else
            txtCustID.Text = CType(value, String)
        End If


    End Sub

    Private Sub txtCustID_TextChanged(sender As Object, e As EventArgs) Handles txtCustID.TextChanged
        'SELECT A.PnCID, A.DateHandedIn, A.PTTime, A.PurchaseAmount, A.ExpiryDate, B.Quantity, B.Description, B.SerialNr, B.QuantitySold
        'FROM (PawningTransactions A INNER JOIN
        'PawnStock B ON A.PnCID = B.PnCID)
        'WHERE (A.CustID = ?) AND (A.Status <> 2)

        Dim table As New DataTable

        Try
            If txtCustID.Text <> "" Then
                com.CommandText = "SELECT A.PnCID, A.DateHandedIn, format(A.PTTime,'HH:MM') as PTTime, A.PurchaseAmount, A.ExpiryDate, B.Quantity, B.Description, B.SerialNr, B.QuantitySold " & _
                                  "FROM (PawningTransactions A INNER JOIN PawnStock B ON A.PnCID = B.PnCID) WHERE (A.Status <> 2) AND (A.CustID = " & txtCustID.Text & ")"
                Dim reader As OleDbDataReader = com.ExecuteReader()
                table.Load(reader)

                With dgPawnInList
                    .DataSource = table
                End With
                'While reader.Read()
                '    Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                '    Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                '    'duplicate entry
                '    Dim intResponse As Integer
                '    intResponse = MsgBox("A customer with similar details already exists." & vbCrLf & vbCrLf & _
                '        "Name: " & reader(1).ToString() & vbCrLf & _
                '        "Customer ID: " & reader(0).ToString() & vbCrLf & _
                '        "RSA ID: " & reader(2).ToString() & vbCrLf & _
                '        "Tel(H): " & reader(3).ToString() & vbCrLf & _
                '        "Tel(W): " & reader(4).ToString() & vbCrLf & _
                '        "Address: " & reader(5).ToString() & vbCrLf & vbCrLf & "Do you still wish to add a new customer?" _
                '        , vbInformation + vbYesNo + vbDefaultButton2, "Duplicate Names")
                'End While
                reader.Close()

                If dgPawnInList.RowCount <> 0 Then
                    dgPawnInList.Visible = True
                    lblNoTransactions.Visible = False
                Else
                    dgPawnInList.Visible = False
                    lblNoTransactions.Visible = True
                End If

            End If

            dgPawnInList.Columns("DateHandedIn").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgPawnInList.Columns("PTTime").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgPawnInList.Columns("Quantity").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgPawnInList.Columns("QuantitySold").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgPawnInList.Columns("ExpiryDate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgPawnInList.Columns("PurchaseAmount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight


            dgPawnInList.AutoResizeColumnHeadersHeight(DataGridViewAutoSizeColumnMode.ColumnHeader)
            dgPawnInList.AutoResizeRowHeadersWidth(DataGridViewAutoSizeColumnMode.ColumnHeader)
            dgPawnInList.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
            dgPawnInList.Refresh()

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try



    End Sub

    Private Sub dgPawnInList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgPawnInList.CellContentClick
        Dim value As Object = dgPawnInList.Rows(e.RowIndex).Cells(0).Value

        If IsDBNull(value) Then
            txtPNCID.Text = ""
        Else
            txtPNCID.Text = CType(value, String)
            cmdDetail_Click(Nothing, Nothing)

        End If


    End Sub
End Class
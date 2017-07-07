Imports System.Data.OleDb

Public Class frmCustMngmnt

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdClearFields_Click(sender As Object, e As EventArgs) Handles cmdClearFields.Click
        ClearAllFields()

    End Sub

    Private Sub cmdSearchCustomer_Click(sender As Object, e As EventArgs) Handles cmdSearchCustomer.Click
        'Try
        '    Me.CustomersTableAdapter.SearchCustByName(Me.CBB_DataSet.Customers, "%" & txtName.Text & "%", _
        '"%" & txtIDnr.Text & "%", "%" & txtTelW.Text & "%", _
        '"%" & txtTelH.Text & "%", "%" & txtAddress.Text & "%")
        'Catch ex As System.Exception
        '    System.Windows.Forms.MessageBox.Show(ex.Message)
        'End Try
        Try
            If txtCustID.Text <> "" Then
                Me.CustomersTableAdapter.SearchCustomer(Me.CBB_DataSet.Customers, CType(txtCustID.Text, Integer))
            Else
                Me.CustomersTableAdapter.SearchCustByName(Me.CBB_DataSet.Customers, "%" & txtName.Text & "%", _
                                                          "%" & txtIDnr.Text & "%", "%" & txtTelW.Text & "%", _
                                                          "%" & txtTelH.Text & "%", "%" & txtAddress.Text & "%")
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

    'Private Sub frmCustMngmnt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    TODO: This line of code loads data into the 'CBB_DataSet.Customers' table. You can move, or remove it, as needed.
    '    Me.CustomersTableAdapter.Fill(Me.CBB_DataSet.Customers)

    'End Sub

    'Private Sub SearchCustomerToolStripButton1_Click(sender As Object, e As EventArgs)
    '    Try
    '        Me.CustomersTableAdapter.SearchCustomer(Me.CBB_DataSet.Customers, CType(CustIDToolStripTextBox.Text, Integer), NameToolStripTextBox.Text, IDNrToolStripTextBox.Text, TelWToolStripTextBox.Text, TelHToolStripTextBox.Text, AddressToolStripTextBox.Text)
    '    Catch ex As System.Exception
    '        System.Windows.Forms.MessageBox.Show(ex.Message)
    '    End Try

    'End Sub

    'Private Sub SearchCustByNameToolStripButton1_Click(sender As Object, e As EventArgs)
    '    Try
    '        Me.CustomersTableAdapter.SearchCustByName(Me.CBB_DataSet.Customers, NameToolStripTextBox2.Text, IDNrToolStripTextBox2.Text, TelWToolStripTextBox2.Text, TelHToolStripTextBox2.Text, AddressToolStripTextBox2.Text)
    '    Catch ex As System.Exception
    '        System.Windows.Forms.MessageBox.Show(ex.Message)
    '    End Try

    'End Sub

    Private Sub dgCustomerList_Click(sender As Object, e As EventArgs) Handles dgCustomerList.Click
        Try
            With dgCustomerList
                txtCustID.Text = .SelectedRows(0).Cells(0).Value.ToString
                txtName.Text = .SelectedRows(0).Cells(1).Value.ToString
                txtIDnr.Text = .SelectedRows(0).Cells(2).Value.ToString
                txtTelW.Text = .SelectedRows(0).Cells(3).Value.ToString
                txtTelH.Text = .SelectedRows(0).Cells(4).Value.ToString
                txtAddress.Text = .SelectedRows(0).Cells(5).Value.ToString
            End With
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub frmCustMngmnt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ClearAllFields()

    End Sub

    Private Sub ClearAllFields()
        'Dim lRows As Long

        'Clear text fields
        txtAddress.Text = ""
        txtCustID.Text = ""
        txtIDnr.Text = ""
        txtName.Text = ""
        txtTelH.Text = ""
        txtTelW.Text = ""
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


    End Sub

    Private Sub cmdUpdateCustomerInformation_Click(sender As Object, e As EventArgs) Handles cmdUpdateCustomerInformation.Click
        Dim intResponse As Integer

        If txtCustID.Text = "" Then
            MsgBox("Customer detail not selected", vbCritical + vbOKOnly)
            Exit Sub
        End If

        intResponse = MsgBox("Do you really want to update the customer information?", vbYesNo + vbDefaultButton2 + vbInformation, "Confirm update")
        If intResponse = vbNo Then Exit Sub

        Try
            'Me.CustomersTableAdapter.Transaction.Begin()
            Me.CustomersTableAdapter.UpdCust(txtName.Text, txtIDnr.Text, txtTelW.Text, txtTelH.Text, txtAddress.Text, CType(txtCustID.Text, Integer))
            'Me.CustomersTableAdapter.Transaction.Commit()

            cmdSearchCustomer_Click(Me, System.EventArgs.Empty)

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try



    End Sub

    Private Sub cmdAddCustomer_Click(sender As Object, e As EventArgs) Handles cmdAddCustomer.Click
        'validate input
        Dim sTmp As String

        sTmp = ""

        If Trim(txtName.Text) = "" Then
            MsgBox("You must specify a name.", vbInformation, "Customer Information")
            Exit Sub
        Else
            'add customer
            'check for duplicate customers
            com.CommandText = "Select CustID, Name, IDnr, TelH, TelW, Address from Customers where Name like '%" & Trim(txtName.Text) & "%' and IDnr like '%" & Trim(txtIDnr.Text) & "%'"
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'duplicate entry
                Dim intResponse As Integer
                intResponse = MsgBox("A customer with similar details already exists." & vbCrLf & vbCrLf & _
                    "Name: " & reader(1).ToString() & vbCrLf & _
                    "Customer ID: " & reader(0).ToString() & vbCrLf & _
                    "RSA ID: " & reader(2).ToString() & vbCrLf & _
                    "Tel(H): " & reader(3).ToString() & vbCrLf & _
                    "Tel(W): " & reader(4).ToString() & vbCrLf & _
                    "Address: " & reader(5).ToString() & vbCrLf & vbCrLf & "Do you still wish to add a new customer?" _
                    , vbInformation + vbYesNo + vbDefaultButton2, "Duplicate Names")
                If intResponse = vbNo Then
                    'cancel new customer
                    reader.Close()
                    sTmp = txtName.Text
                    ClearAllFields()
                    txtName.Text = sTmp
                    cmdSearchCustomer_Click(Me, System.EventArgs.Empty)
                    Exit Sub
                End If
            End While
            reader.Close()

            Try
                'Me.CustomersTableAdapter.Transaction.Begin()
                Me.CustomersTableAdapter.NewCustomer(txtName.Text, txtIDnr.Text, txtTelW.Text, txtTelH.Text, txtAddress.Text)

                com.CommandText = "SELECT MAX(CustID) as LASTCustID FROM Customers"
                reader = com.ExecuteReader()
                While reader.Read()
                    MsgBox("CustomerID of new customer is: " & reader(0).ToString(), vbInformation, "New Customer ID")
                End While
                reader.Close()
                cmdSearchCustomer_Click(Me, System.EventArgs.Empty)
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            End Try



        End If
    End Sub

    Private Sub cmdUseThisCustomer_Click(sender As Object, e As EventArgs) Handles cmdUseThisCustomer.Click
        If txtCustID.Text <> "" Then
            With frmPawnOut
                .txtCustID.Text = txtCustID.Text
                .lblAdres.Text = txtAddress.Text
                .lblName.Text = txtName.Text
                .lblTelH.Text = txtTelH.Text
                .lblTelW.Text = txtTelW.Text
                .lblID.Text = txtIDnr.Text
            End With
            Me.Close()
        Else
            MsgBox("Please select Customer!", vbInformation + vbOKOnly)
        End If

        
    End Sub

    Private Sub cmdShowHistory_Click(sender As Object, e As EventArgs) Handles cmdShowHistory.Click
        If txtCustID.Text <> "" Then
            With frmCustomerHistory
                .txtCustID.Text = Me.txtCustID.Text
                .txtCustName.Text = Me.txtName.Text
                .ShowDialog(Me)
            End With
        Else
            MsgBox("Please select Customer!", vbInformation + vbOKOnly)
        End If

    End Sub
End Class
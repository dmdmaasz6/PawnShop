Imports System.Data.OleDb

Public Class frmStockPurchase

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub frmStockPurchase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        grpNormalStock.Left = 15
        grpStandardStock.Left = 15
        grpTangoStock.Left = 15

        rbNormal.Checked = True
        rbNormal.Font = New System.Drawing.Font(rbNormal.Font, FontStyle.Bold)
        grpNormalStock.Visible = True
        rbStandard.Checked = False
        rbStandard.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Regular)
        grpStandardStock.Visible = False
        rbTango.Checked = False
        rbTango.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Regular)
        grpTangoStock.Visible = False

        Clearfileds()

    End Sub

    Private Sub rbNormal_CheckedChanged(sender As Object, e As EventArgs) Handles rbNormal.CheckedChanged
        rbNormal.Font = New System.Drawing.Font(rbNormal.Font, FontStyle.Bold)
        rbStandard.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Regular)
        rbTango.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Regular)
        grpNormalStock.Visible = True
        grpStandardStock.Visible = False
        grpTangoStock.Visible = False

        If rbNormal.Checked = True Then
            Clearfileds()
        End If

    End Sub

    Private Sub rbStandard_CheckedChanged(sender As Object, e As EventArgs) Handles rbStandard.CheckedChanged
        rbNormal.Font = New System.Drawing.Font(rbNormal.Font, FontStyle.Regular)
        rbStandard.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Bold)
        rbTango.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Regular)
        grpNormalStock.Visible = False
        grpStandardStock.Visible = True
        grpTangoStock.Visible = False

        If rbStandard.Checked = True Then
            Clearfileds()
            LoadStockList()
        End If

    End Sub

    Private Sub rbTango_CheckedChanged(sender As Object, e As EventArgs) Handles rbTango.CheckedChanged
        rbNormal.Font = New System.Drawing.Font(rbNormal.Font, FontStyle.Regular)
        rbStandard.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Regular)
        rbTango.Font = New System.Drawing.Font(rbTango.Font, FontStyle.Bold)
        grpTangoStock.Visible = True
        grpNormalStock.Visible = False
        grpStandardStock.Visible = False

        If rbTango.Checked = True Then
            Clearfileds()
            LoadTangoList()
        End If

    End Sub

    Private Sub Clearfileds()
        'Normal Stock
        txtName.Text = ""
        txtIDNr.Text = ""
        txtAddress.Text = ""
        txtDescription.Text = ""
        txtNSID.Text = ""

        'Standard Stock
        Try
            dgStdStockManagment.ClearSelection()
            If dgStdStockManagment.Rows.Count <> 0 Then
                Do While dgStdStockManagment.Rows.Count <> 0
                    dgStdStockManagment.Rows.Remove(dgStdStockManagment.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

        'Tango Stock
        Try
            dgTangoStockManagment.ClearSelection()
            If dgTangoStockManagment.Rows.Count <> 0 Then
                Do While dgTangoStockManagment.Rows.Count <> 0
                    dgTangoStockManagment.Rows.Remove(dgTangoStockManagment.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

        'Transaction
        txtQuantity.Text = "1"
        txtRecommendedSellingPrice.Text = ""
        txtPurchasePrice.Text = ""

    End Sub

    Private Sub cmdPurchase_Click(sender As Object, e As EventArgs) Handles cmdPurchase.Click
        Dim sSQL As String
        Dim CurrNSID As String
        CurrNSID = ""

        'Input Validation
        If rbStandard.Checked = True Then ' Standard Stock
            If txtNSID.Text = "" Then
                MissingInfo("a stock")
                Exit Sub
            End If
        ElseIf rbTango.Checked = True Then 'Tango Stock
            If txtNSID.Text = "" Then
                MissingInfo("a stock")
                Exit Sub
            End If
        Else    'Normal Stock
            If Trim(txtName.Text) = "" Then
                MissingInfo("a name for the supplier")
                txtName.Focus()
                Exit Sub
            ElseIf Trim(txtDescription.Text) = "" Then
                MissingInfo("a description for the stock item")
                txtDescription.Focus()
                Exit Sub
            End If
        End If
        ' Normal and Standard Stock
        If Trim(txtPurchasePrice.Text) = "" Then
            MissingInfo("a purchase price for the stock item")
            txtPurchasePrice.Focus()
            Exit Sub
        ElseIf Trim(txtRecommendedSellingPrice.Text) = "" Then
            MissingInfo("a recommended selling price for the stock item")
            txtRecommendedSellingPrice.Focus()
            Exit Sub
        End If

        Try
            ' Insert stock purchase into NormalStock table
            If rbStandard.Checked = True Then 'Standard Stock
                'Update Quantity and RecommendedSellingPrice
                '            cmd.CommandText = "Update NormalStock set RecommendedSellingPrice = " & txtRecommendedSellingPrice.Text & ", Quantity = Quantity + " & txtQuantity.Text & " where NSID = " & dePNC.rscmdStandardStockCodeList.Fields("NSID").Value
                sSQL = "Update NormalStock set RecommendedSellingPrice = " & txtRecommendedSellingPrice.Text & ", Quantity = Quantity + " & txtQuantity.Text & " where NSID = " & txtNSID.Text
                com.CommandText = sSQL
                Dim reader As OleDbDataReader = com.ExecuteReader()
                reader.Close()
            ElseIf rbTango.Checked = True Then 'Tango Stock
                'Update Quantity and RecommendedSellingPrice
                '            cmd.CommandText = "Update NormalStock set RecommendedSellingPrice = " & txtRecommendedSellingPrice.Text & ", Quantity = Quantity + " & txtQuantity.Text & " where NSID = " & dePNC.rscmdTangoStockCodeList.Fields("NSID").Value
                sSQL = "Update NormalStock set RecommendedSellingPrice = " & txtRecommendedSellingPrice.Text & ", Quantity = Quantity + " & txtQuantity.Text & " where NSID = " & txtNSID.Text
                com.CommandText = sSQL
                Dim reader As OleDbDataReader = com.ExecuteReader()
                reader.Close()
            Else 'Normal Stock
                Me.NormalStockTableAdapter1.Insert(FormatDateTime(dtpDate.Value, DateFormat.ShortDate), txtName.Text, txtIDNr.Text, txtAddress.Text, txtDescription.Text, CInt(txtQuantity.Text), 0, 0, FormatCurrency(txtPurchasePrice.Text), FormatCurrency(txtRecommendedSellingPrice.Text), 0)
            End If

            ' Retrieve NSID from lastly added item and display
            'Normal Stock
            If rbNormal.Checked = True Then 'Normal Stock
                sSQL = "Select max(NSID) as MaxNSID from NormalStock" ', dePNC.cnPNC"
                com.CommandText = sSQL
                Dim reader As OleDbDataReader = com.ExecuteReader()
                While reader.Read()
                    Console.WriteLine(reader(0).ToString())
                    Debug.WriteLine(reader(0).ToString())
                    CurrNSID = reader(0).ToString()
                End While
                reader.Close()
                MsgBox("Mark this stock item(s) with the following Normal Stock ID (NSID): " & CurrNSID, vbInformation, "Mark Stock Items")

            End If

            ' Insert Stock Purchase into CashTransactions table
            If rbStandard.Checked = True Then 'Standard Stock
                ' Standard Stock Purchase
                'dePNC.cmdInsertCashTransaction(gstrCR, "Standard Stock Purchase", dtpDate.Value, 0, 0, txtQuantity.Text, dePNC.rscmdStandardStockCodeList.Fields("NSID").Value, 0, dePNC.rscmdStandardStockCodeList.Fields("Description").Value, -((CCur(txtPurchasePrice.Text)) * txtQuantity.Text), txtPurchasePrice.Text, 0, 0)
                Me.CashTransactionsTableAdapter1.Insert("Standard Stock Purchase", FormatDateTime(dtpDate.Value, DateFormat.ShortDate), 0, 0, txtQuantity.Text, txtNSID.Text, 0, txtDescription.Text, -((FormatCurrency(txtPurchasePrice.Text)) * txtQuantity.Text), 0, 0, FormatCurrency(txtPurchasePrice.Text), p_CashRegister)
                If p_Audit = True Then AuditToFile("Standard Stock Purchase:", "Date: " & FormatDateTime(dtpDate.Value, DateFormat.ShortDate) & ", NSID: " & txtNSID.Text & ", " & txtDescription.Text & ", Quantity: " & txtQuantity.Text & ", Purchase Amount: " & (FormatCurrency(txtPurchasePrice.Text)) & ", Recommended Selling Price: " & (FormatCurrency(txtRecommendedSellingPrice.Text)))
            ElseIf rbTango.Checked = True Then 'Tango Stock
                ' Standard Stock Purchase
                'dePNC.cmdInsertCashTransaction(gstrCR, "Tango Stock Purchase", dtpDate.Value, 0, 0, txtQuantity.Text, dePNC.rscmdTangoStockCodeList.Fields("NSID").Value, 0, dePNC.rscmdTangoStockCodeList.Fields("Description").Value, -((CCur(txtPurchasePrice.Text)) * txtQuantity.Text), txtPurchasePrice.Text, 0, 0)
                Me.CashTransactionsTableAdapter1.Insert("Tango Stock Purchase", FormatDateTime(dtpDate.Value, DateFormat.ShortDate), 0, 0, txtQuantity.Text, txtNSID.Text, 0, txtDescription.Text, -((FormatCurrency(txtPurchasePrice.Text)) * txtQuantity.Text), 0, 0, FormatCurrency(txtPurchasePrice.Text), p_CashRegister)
                If p_Audit = True Then AuditToFile("Tango Stock Purchase:", "Date: " & FormatDateTime(dtpDate.Value, DateFormat.ShortDate) & ", NSID: " & txtNSID.Text & ", " & txtDescription.Text & ", Quantity: " & txtQuantity.Text & ", Purchase Amount: " & (FormatCurrency(txtPurchasePrice.Text)) & ", Recommended Selling Price: " & (FormatCurrency(txtRecommendedSellingPrice.Text)))
            Else
                ' Normal Stock Purchase
                Me.CashTransactionsTableAdapter1.Insert("Normal Stock Purchase", FormatDateTime(dtpDate.Value, DateFormat.ShortDate), 0, 0, txtQuantity.Text, CurrNSID, 0, txtDescription.Text, -((FormatCurrency(txtPurchasePrice.Text)) * txtQuantity.Text), 0, 0, FormatCurrency(txtPurchasePrice.Text), p_CashRegister)
                If p_Audit = True Then AuditToFile("Normal Stock Purchase:", "Date: " & FormatDateTime(dtpDate.Value, DateFormat.ShortDate) & ", " & txtDescription.Text & ", Quantity: " & txtQuantity.Text & ", Purchase Amount: " & (FormatCurrency(txtPurchasePrice.Text)) & ", Recommended Selling Price: " & (FormatCurrency(txtRecommendedSellingPrice.Text)))
            End If

            Me.Close()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtPurchasePrice_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPurchasePrice.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Price must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If

    End Sub

    Private Sub txtRecommendedSellingPrice_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtRecommendedSellingPrice.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Price must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If

    End Sub

    Private Sub txtQuanity_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuantity.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Quantity must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If

    End Sub

    Private Sub txtPurchasePrice_LostFocus(sender As Object, e As EventArgs) Handles txtPurchasePrice.LostFocus
        SetCurr(txtPurchasePrice)
    End Sub

    Private Sub txtRecommendedSellingPrice_LostFocus(sender As Object, e As EventArgs) Handles txtRecommendedSellingPrice.LostFocus
        SetCurr(txtRecommendedSellingPrice)
    End Sub

    Private Sub LoadStockList()
        Try
            dgStdStockManagment.ClearSelection()

            With Me.dgStdStockManagment
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            com.CommandText = "SELECT NSID, Description, RecommendedSellingPrice, (Quantity - QuantitySold - QuantityLayBuy) as QuantityAvailable, Quantity, QuantitySold, QuantityLayBuy, CellphoneAssessories FROM NormalStock WHERE (CellphoneAssessories = 1) order by Description"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Me.dgStdStockManagment.Rows.Add(reader(0).ToString(), reader(1).ToString(), GetValue(reader(2).ToString()))
            End While
            reader.Close()

            Me.dgStdStockManagment.ClearSelection()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub LoadTangoList()
        Try
            dgTangoStockManagment.ClearSelection()

            With Me.dgTangoStockManagment
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            com.CommandText = "SELECT NSID, Description, RecommendedSellingPrice, (Quantity - QuantitySold - QuantityLayBuy) as QuantityAvailable, Quantity, QuantitySold, QuantityLayBuy, CellphoneAssessories FROM NormalStock WHERE (CellphoneAssessories = 2) order by Description"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Me.dgTangoStockManagment.Rows.Add(reader(0).ToString(), reader(1).ToString(), GetValue(reader(2).ToString()))
            End While
            reader.Close()

            Me.dgTangoStockManagment.ClearSelection()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Function GetValue(sValIn As String) As String
        Dim sTmp As String
        sTmp = FormatNumber(sValIn, 2)
        If Microsoft.VisualBasic.Left(sValIn, 1) = "-" Then
            GetValue = Replace(sTmp, "-", "-" & p_Currency)
        Else
            GetValue = p_Currency & sTmp
        End If

    End Function

    Private Sub dgStdStockManagment_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgStdStockManagment.CellMouseClick
        Dim sRecSell As String

        txtNSID.Text = Me.dgStdStockManagment.Item(0, dgStdStockManagment.CurrentRow.Index).Value
        txtDescription.Text = Me.dgStdStockManagment.Item(1, dgStdStockManagment.CurrentRow.Index).Value
        sRecSell = Trim(Me.dgStdStockManagment.Item(2, dgStdStockManagment.CurrentRow.Index).Value)
        txtRecommendedSellingPrice.Text = Microsoft.VisualBasic.Right(sRecSell, Len(sRecSell) - 1)

    End Sub

    Private Sub dgTangoStockManagment_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgTangoStockManagment.CellMouseClick
        Dim sRecSell As String

        txtNSID.Text = Me.dgTangoStockManagment.Item(0, dgTangoStockManagment.CurrentRow.Index).Value
        txtDescription.Text = Me.dgTangoStockManagment.Item(1, dgTangoStockManagment.CurrentRow.Index).Value
        sRecSell = Trim(Me.dgTangoStockManagment.Item(2, dgTangoStockManagment.CurrentRow.Index).Value)
        txtRecommendedSellingPrice.Text = Microsoft.VisualBasic.Right(sRecSell, Len(sRecSell) - 1)

    End Sub

End Class
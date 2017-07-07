Imports System.Data.OleDb

Public Class frmPawnOut
    Dim sSQL As String
    Private lngPNC As Long 'PNC for next pawning out
    Private strMonthNr As String 'MonthNr for next pawning out

    Public Sub ClearForm()
        lblAdres.Text = ""
        lblName.Text = ""
        lblTelH.Text = ""
        lblTelW.Text = ""
        lblID.Text = ""

        txtQty1.Text = "1"
        txtQty2.Text = "1"
        txtQty3.Text = "1"
        txtQty4.Text = "1"

        txtDescr1.Text = ""
        txtDescr2.Text = ""
        txtDescr3.Text = ""
        txtDescr4.Text = ""

        txtSrNR1.Text = ""
        txtSrNR2.Text = ""
        txtSrNR3.Text = ""
        txtSrNR4.Text = ""

        txtBuyBackAmount.Text = "0"
        txtPurchaseAmount.Text = "0"

        dtpDateHandedIn.Value = Now
        dtpExpiryDate.Value = Now

    End Sub

    Private Sub frmPawnOut_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblPTTime.Text = FormatDateTime(Now, DateFormat.ShortTime)
        txtCustID.Text = ""
        ClearForm()

    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        txtCustID.Text = ""
        ClearForm()
        Me.Close()

    End Sub

    Private Sub cmdSearch_Click(sender As Object, e As EventArgs) Handles cmdSearch.Click
        txtCustID.Text = ""
        ClearForm()
        With frmCustMngmnt
            .cmdAddCustomer.Enabled = False
            .cmdUpdateCustomerInformation.Enabled = False
            .cmdShowHistory.Visible = False
            .cmdUseThisCustomer.Visible = True
            .ShowDialog(Me)
        End With
    End Sub

    Private Sub cmdDetail_Click(sender As Object, e As EventArgs) Handles cmdDetail.Click
        Dim bFound As Boolean

        Try
            bFound = False
            ClearForm()

            If Trim(txtCustID.Text) = "" Then
                MsgBox("You must specify a valid CustomerID", vbCritical, "Customer ID")
                txtCustID.Text = ""
                txtCustID.Focus()
            Else

                sSQL = "SELECT NAME, IDnr, TelW, TelH, Address FROM Customers WHERE CustID = " & txtCustID.Text
                com.CommandText = sSQL
                Dim reader As OleDbDataReader = com.ExecuteReader()
                While reader.Read()
                    lblName.Text = reader(0).ToString
                    lblID.Text = reader(1).ToString
                    lblTelW.Text = reader(2).ToString
                    lblTelH.Text = reader(3).ToString
                    lblAdres.Text = reader(4).ToString
                    bFound = True
                End While
                reader.Close()

                If bFound = False Then
                    MsgBox("CustomerID " & txtCustID.Text & " does not exist.", vbCritical, "Invalid CustomerID")
                    txtCustID.Focus()
                End If

            End If
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Function ReadyToPawn() As Boolean

        If Trim(txtCustID.Text) = "" Then
            MissingInfo("a Customer ID")
            txtCustID.Focus()
            ReadyToPawn = False
        ElseIf Trim(txtDescr1.Text) = "" Then
            MissingInfo("at least one pawn item")
            txtDescr1.Focus()
            ReadyToPawn = False
        ElseIf CDbl(txtPurchaseAmount.Text) <= 0 Then
            MissingInfo("a Purchase Price")
            txtPurchaseAmount.Focus()
            ReadyToPawn = False
        ElseIf CDbl(txtBuyBackAmount.Text) <= 0 Then
            MissingInfo("a Buy Back Price")
            txtBuyBackAmount.Focus()
            ReadyToPawn = False
        ElseIf dtpDateHandedIn.Value > dtpExpiryDate.Value Then
            MsgBox("Expiry Date must be set a date later that the Date Handed in.", vbInformation, "Date Selection")
            ReadyToPawn = False
        ElseIf CDec(txtPurchaseAmount.Text) > CDec(txtBuyBackAmount.Text) Then
            MsgBox("Purchase Amount cannot exceed Buy Back Amount")
            ReadyToPawn = False
        Else
            ReadyToPawn = True
        End If

    End Function

    Private Function NextPnC() As Long
        Dim PnCID As Long

        Try
            sSQL = "SELECT MAX(PnCID) AS LastPNC FROM PawningTransactions"

            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                PnCID = CLng(reader(0).ToString)
            End While
            reader.Close()

            NextPnC = PnCID + 1

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Function

    Private Function NextMonthNr(lngPNC As Long) As String
        'Get latest MonthNr
        Dim strLastMonthNr As String

        Try

            sSQL = "SELECT MonthNr FROM PawningTransactions WHERE pncid = " & lngPNC
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                strLastMonthNr = reader(0).ToString
            End While
            reader.Close()

            'Get month portion
            Dim intFS1 As Integer
            Dim intFS2 As Integer
            Dim intMonth As Integer 'month portion of last MonthNr
            intFS1 = InStr(1, strLastMonthNr, "/")
            intFS2 = InStr(intFS1 + 1, strLastMonthNr, "/")
            intMonth = Mid(strLastMonthNr, intFS1 + 1, intFS2 - intFS1 - 1)
            'Determine if still in same month as intMonth
            If intMonth = Month(Now) Then
                ' still the same month
                Dim strDay As String
                strDay = Mid(strLastMonthNr, 1, intFS1 - 1)
                NextMonthNr = strDay + 1 & "/" & Month(Now) & "/" & Year(Now)
            Else
                ' beginning of the next month
                NextMonthNr = "1/" & Month(Now) & "/" & Year(Now)
            End If

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Function

    Public Function FinalExpiryDate(dtmExpire As Date) As Date
        ' Add 7 days, excluding Sundays
        Dim i As Integer
        For i = 1 To 7 ' From day of transaction to day of available to sell
            dtmExpire = DateAdd("w", 1, dtmExpire)
            '        If Weekday(dtmExpire) = 1 Then 'Sunday
            '            ' Sunday, therefore ignore increment
            '            i = i - 1
            '        End If
        Next i
        FinalExpiryDate = dtmExpire
    End Function

    Private Sub cmdPawnOut_Click(sender As Object, e As EventArgs) Handles cmdPawnOut.Click

        Try
            ClearStockList()

            sAgreeTime = Replace(FormatDateTime(Now, DateFormat.LongTime), ":", "")

            'Prepare PNCID and MonthNr
            lngPNC = NextPnC()
            strMonthNr = NextMonthNr(lngPNC - 1)

            If ReadyToPawn() = True Then
                ' add entry to PawningTransactions table
                InsertNewPawningTransactions(lngPNC, strMonthNr, CDate(lblPTTime.Text), CLng(txtCustID.Text), CDbl(txtPurchaseAmount.Text), CDate(dtpDateHandedIn.Value), CDbl(txtBuyBackAmount.Text), CDate(dtpExpiryDate.Value), FinalExpiryDate(dtpExpiryDate.Value))
                ' Insert into CashTransactions
                InsertCashTransaction(lngPNC, CDate(dtpDateHandedIn.Value), txtPurchaseAmount.Text)


                arrStockSold(iStock, 0) = "Quantity"
                arrStockSold(iStock, 1) = "Description"
                arrStockSold(iStock, 2) = "Serial Nr:"
                iStock = iStock + 1

                ' check whether descriptions are given
                If Trim(txtDescr1.Text) <> "" Then
                    InsertNewPawnStock(lngPNC, txtDescr1.Text, txtSrNR1.Text, CInt(txtQty1.Text))
                    arrStockSold(iStock, 0) = (Trim(txtQty1.Text))
                    arrStockSold(iStock, 1) = (Trim(txtDescr1.Text))
                    arrStockSold(iStock, 2) = (Trim(txtSrNR1.Text))
                    iStock = iStock + 1
                End If

                If Trim(txtDescr2.Text) <> "" Then
                    InsertNewPawnStock(lngPNC, txtDescr2.Text, txtSrNR2.Text, CInt(txtQty2.Text))
                    arrStockSold(iStock, 0) = (Trim(txtQty2.Text))
                    arrStockSold(iStock, 1) = (Trim(txtDescr2.Text))
                    arrStockSold(iStock, 2) = (Trim(txtSrNR2.Text))
                    iStock = iStock + 1
                End If

                If Trim(txtDescr3.Text) <> "" Then
                    InsertNewPawnStock(lngPNC, txtDescr3.Text, txtSrNR3.Text, CInt(txtQty3.Text))
                    arrStockSold(iStock, 0) = (Trim(txtQty3.Text))
                    arrStockSold(iStock, 1) = (Trim(txtDescr3.Text))
                    arrStockSold(iStock, 2) = (Trim(txtSrNR3.Text))
                    iStock = iStock + 1
                End If

                If Trim(txtDescr4.Text) <> "" Then
                    InsertNewPawnStock(lngPNC, txtDescr4.Text, txtSrNR4.Text, CInt(txtQty4.Text))
                    arrStockSold(iStock, 0) = (Trim(txtQty4.Text))
                    arrStockSold(iStock, 1) = (Trim(txtDescr4.Text))
                    arrStockSold(iStock, 2) = (Trim(txtSrNR4.Text))
                    iStock = iStock + 1
                End If

                MsgBox("Information updated to database", vbInformation, "Database updated")

                Template_Agreement(CStr(lngPNC), Trim(strMonthNr), Trim(lblPTTime.Text), _
                                   Trim(txtCustID.Text), Trim(lblName.Text), Trim(lblID.Text), _
                                   Trim(lblTelW.Text), Trim(lblTelH.Text), Trim(lblAdres.Text), _
                                   Trim(txtPurchaseAmount.Text), Trim(txtBuyBackAmount.Text), _
                                   FormatDateTime(dtpDateHandedIn.Value, DateFormat.LongDate), _
                                   FormatDateTime(dtpExpiryDate.Value, DateFormat.LongDate))
                'Format((Trim(dtpDateHandedIn.Value)), "DD-MMM-YYYY"), _
                'Format((Trim(dtpExpiryDate.Value)), "DD-MMM-YYYY"))

                Template_AgreementTearOff(CStr(lngPNC), _
                          Trim(lblName.Text), _
                          Trim(txtBuyBackAmount.Text), _
                          FormatDateTime(dtpExpiryDate.Value, DateFormat.LongDate), _
                          Trim(txtCustID.Text))
                'Format((Trim(dtpExpiryDate.Value)), "DD-MMM-YYYY"))

            Else
                Exit Sub
            End If

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

    Private Sub txtPurchaseAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPurchaseAmount.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Purchase Amount must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If
    End Sub

    Private Sub txtBuyBackAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtBuyBackAmount.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Buyback Amount must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If
    End Sub

    Private Sub txtPurchaseAmount_LostFocus(sender As Object, e As EventArgs) Handles txtPurchaseAmount.LostFocus
        SetCurr(txtPurchaseAmount)

    End Sub

    Private Sub txtBuyBackAmount_LostFocus(sender As Object, e As EventArgs) Handles txtBuyBackAmount.LostFocus
        SetCurr(txtBuyBackAmount)

    End Sub


    Private Sub InsertNewPawnStock(lPnCID As Long, sDescription As String, sSerialNumber As String, sQuantity As String)
        Dim sLastCmd As String

        Try
            sLastCmd = ""

            'Insert into PawnStock
            sSQL = "INSERT INTO PawnStock " & _
                "(PnCID, Description, SerialNr, Status, Quantity) VALUES " & _
                "(" & lPnCID & ", '" & sDescription & "', '" & sSerialNumber & "', 1, '" & sQuantity & "')"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            reader2.Close()

            sLastCmd = "cmdNewPawnStock" & "," & lPnCID & ", " & sDescription & ", " & sSerialNumber & ", 1, " & sQuantity
            'MsgBox(sLastCmd)

            If p_Audit = True Then AuditToFile("New Pawn Stock", sLastCmd)

        Catch ex As Exception
            MsgBox("An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
                "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
                "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
                "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
                "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update.")
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub InsertNewPawningTransactions(lngPncID As Long, sMonthNr As String, dtDTTime As Date, lCustomerID As Long, dPurchaseAmount As Double, _
                                             dtDateHandedIn As Date, dBuyBackAmount As Double, dtExpiryDate As Date, dtFinalExpiryDate As Date)
        Dim sLastCmd As String

        Try
            sLastCmd = ""

            'Insert into PawnStock
            ''dePNC.cmdNewPawningTransaction(lngPNC, strMonthNr, CDate(lblPTTime.Caption), CLng(txtCustomerID.Text), CDbl(txtPurchaseAmount.Text), CDate(dtpDateHandedIn.Value), CDbl(txtBuyBackAmount.Text), CDate(dtpExpiryDate.Value), FinalExpiryDate(dtpExpiryDate.Value))
            sSQL = "INSERT INTO PawningTransactions " & _
                "(PnCID, MonthNr, PTTime, CustID, PurchaseAmount, DateHandedIn, BuyBackAmount, ExpiryDate, Status, ExpiryDateAfter7WorkDays) VALUES " & _
                "(" & lngPncID & ", '" & sMonthNr & "', '" & dtDTTime & "', " & lCustomerID & ", " & dPurchaseAmount & ", '" & dtDateHandedIn & "', " & dBuyBackAmount & ", '" & dtExpiryDate & "', 1, '" & dtFinalExpiryDate & "')"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            reader2.Close()

            sLastCmd = "cmdNewPawningTransactions" & "," & lngPncID & ", " & sMonthNr & ", " & dtDTTime & ", " & lCustomerID & ", " & dPurchaseAmount & ", " & dtDateHandedIn & ", " & dBuyBackAmount & ", " & dtExpiryDate & ", 1, " & dtFinalExpiryDate
            'MsgBox(sLastCmd)

            If p_Audit = True Then AuditToFile("New Pawn Stock", sLastCmd)

        Catch ex As Exception
            MsgBox("An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
                "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
                "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
                "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
                "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update.")
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub InsertCashTransaction(lPncID as long, dtpDateHandedIn as date, sPurchaseAmount as string)
        Dim sLastCmd As String

        Try
            sLastCmd = ""

            'Insert into PawnStock
            ''dePNC.cmdInsertCashTransaction(gstrCR, "Pawning Out", dtpDateHandedIn.Value, lngPNC, 0, 0, 0, 0, "", -(txtPurchaseAmount.Text), 0, 0, 0)
            sSQL = "INSERT INTO CashTransactions " & _
                   "(CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES " & _
                   "('" & p_CashRegister & "','Pawning Out','" & dtpDateHandedIn & "', " & lPncID & ", 0, 0, 0, 0,'','-" & sPurchaseAmount & "',0,0,0)"

            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            reader2.Close()

            sLastCmd = "cmdInsertCashTransaction1" & "," & p_CashRegister & "," & "Pawning Out" & "," & dtpDateHandedIn & ", " & lPncID & ", 0, 0, 0, 0,'','-" & sPurchaseAmount & "',0,0,0)"
            'MsgBox(sLastCmd)

            If p_Audit = True Then AuditToFile("New Pawn Stock", sLastCmd)

        Catch ex As Exception
            MsgBox("An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
                "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
                "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
                "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
                "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update.")
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtPurchaseAmount_TextChanged(sender As Object, e As EventArgs) Handles txtPurchaseAmount.TextChanged

    End Sub
End Class
Imports System.Data.OleDb

Public Class frmTangoStockManagment
    'Public reader As OleDbDataReader = com.ExecuteReader()

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        ClearForm()
        txtPriceUpd.Text = ""
        txtDescriptionUpd.Text = ""
        dgTangoStockManagment.ClearSelection()

        Me.Close()

    End Sub

    Private Sub ClearForm()
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

        txtPrice.Text = ""
        txtDescription.Text = ""
        txtPriceUpd.Text = ""
        txtDescriptionUpd.Text = ""

    End Sub

    Private Sub LoadList()
        Try
            dgTangoStockManagment.ClearSelection()

            With Me.dgTangoStockManagment
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            com.CommandText = "SELECT NSID, Description, RecommendedSellingPrice, (Quantity - QuantitySold - QuantityLayBuy) as QuantityAvailable, Quantity, QuantitySold, QuantityLayBuy, CellphoneAssessories FROM NormalStock WHERE (CellphoneAssessories = 2) order by Description"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Me.dgTangoStockManagment.Rows.Add(reader(0).ToString(), reader(1).ToString(), GetValue(reader(2).ToString()), reader(3).ToString(), _
                                          reader(6).ToString())
            End While
            reader.Close()

            txtDescriptionUpd.Text = ""
            txtPriceUpd.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub frmTangoStockManagment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadList()
        dgTangoStockManagment.ClearSelection()
        txtDescriptionUpd.Text = ""
        txtPriceUpd.Text = ""

    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        Dim intResponse As Integer
        Dim sNSID As String
        Dim sDescription As String

        Try

            sNSID = Me.dgTangoStockManagment.Item(0, dgTangoStockManagment.CurrentRow.Index).Value
            sDescription = Me.dgTangoStockManagment.Item(1, dgTangoStockManagment.CurrentRow.Index).Value

            intResponse = MsgBox("Do you really want to delete the entry?" & vbCrLf & sDescription, MsgBoxStyle.Question + vbYesNo, "Confirm Cancellation")
            If intResponse = vbYes Then

                com.CommandText = "delete from NormalStock where NSID = " & sNSID & " and Description = '" & sDescription & "'"
                Dim reader As OleDbDataReader = com.ExecuteReader()

                reader.Close()

                If p_Audit = True Then AuditToFile("Normal Stock Delete:", com.CommandText)

                ClearForm()
                LoadList()

            End If

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtPrice_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrice.KeyPress
        'If Not (IsNumeric(sender.ToString) Or KeyAscii = 8 Or KeyAscii = 46) Then
        '    MsgBox("Price must have a numeric value.", vbInformation, "Key invalid")
        '    KeyAscii = 0
        'End If
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Price must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If

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

    Private Sub txtPrice_LostFocus(sender As Object, e As EventArgs) Handles txtPrice.LostFocus
        txtPrice.Text = txtPrice.Text & ".00"

    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click

        Try
            If Trim(txtDescription.Text) = "" Then
                MsgBox("No Description supplied!", MsgBoxStyle.Exclamation + vbOKOnly)
                Exit Sub
            End If

            If Len(txtPrice.Text) < 4 Then
                MsgBox("No Selling Price supplied!", MsgBoxStyle.Exclamation + vbOKOnly)
                Exit Sub
            End If

            If Len(txtDescription.Text) > 50 Then
                MsgBox("Description length exceed maximum 50 caharacters allowed!", vbInformation + vbOKOnly, "Description Error")
                Exit Sub
            End If

            com.CommandText = "INSERT INTO NormalStock (Description,RecommendedSellingPrice,CellphoneAssessories,Quantity,QuantitySold,QuantityLayBuy) VALUES " & _
                              "('" & txtDescription.Text & "'," & Microsoft.VisualBasic.Left(txtPrice.Text, Len(txtPrice.Text) - 3) & ",2,0,0,0)"

            Dim reader As OleDbDataReader = com.ExecuteReader()

            reader.Close()

            If p_Audit = True Then AuditToFile("Normal Stock Add:", com.CommandText)

            ClearForm()
            LoadList()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        Dim sNSID As String
        Dim sDescription As String
        Dim sPrice As String

        Try

            If Trim(txtDescriptionUpd.Text) = "" Then
                MsgBox("No Description supplied!", MsgBoxStyle.Exclamation + vbOKOnly)
                Exit Sub
            End If

            If Len(txtPriceUpd.Text) < 4 Then
                MsgBox("No Selling Price supplied!", MsgBoxStyle.Exclamation + vbOKOnly)
                Exit Sub
            End If

            If Len(txtDescriptionUpd.Text) > 50 Then
                MsgBox("Description length exceed maximum 50 caharacters allowed!", vbInformation + vbOKOnly, "Description Error")
                Exit Sub
            End If

            sNSID = Me.dgTangoStockManagment.Item(0, dgTangoStockManagment.CurrentRow.Index).Value
            sDescription = Trim(txtDescriptionUpd.Text)
            sPrice = Microsoft.VisualBasic.Left(txtPriceUpd.Text, Len(txtPriceUpd.Text) - 3)

            If sDescription = Me.dgTangoStockManagment.Item(1, dgTangoStockManagment.CurrentRow.Index).Value Then
                If p_Currency & sPrice & ".00" = Me.dgTangoStockManagment.Item(2, dgTangoStockManagment.CurrentRow.Index).Value Then
                    MsgBox("No new values supplied!" & vbCrLf & "Nothing to update!", MsgBoxStyle.Exclamation + vbOKOnly)
                    Exit Sub
                End If
            End If

            com.CommandText = "Update NormalStock SET Description = '" & sDescription & "',RecommendedSellingPrice = " & sPrice & " WHERE NSID = " & sNSID
            Dim reader As OleDbDataReader = com.ExecuteReader()

            reader.Close()

            If p_Audit = True Then AuditToFile("Normal Stock Update:", com.CommandText)

            ClearForm()
            LoadList()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub dgTangoStockManagment_SelectionChanged(sender As Object, e As EventArgs) Handles dgTangoStockManagment.SelectionChanged
        Try
            txtPriceUpd.Text = Microsoft.VisualBasic.Right(Me.dgTangoStockManagment.Item(2, dgTangoStockManagment.CurrentRow.Index).Value, Len(Me.dgTangoStockManagment.Item(2, dgTangoStockManagment.CurrentRow.Index).Value) - 1)
            txtDescriptionUpd.Text = Me.dgTangoStockManagment.Item(1, dgTangoStockManagment.CurrentRow.Index).Value
        Catch
        End Try
    End Sub

    Private Sub txtPriceUpd_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPriceUpd.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
                MsgBox("Price must have a numeric value.", vbInformation + vbOKOnly, "Key invalid")
            End If
        End If

    End Sub

    Private Sub txtPriceUpd_LostFocus(sender As Object, e As EventArgs) Handles txtPriceUpd.LostFocus
        txtPriceUpd.Text = txtPriceUpd.Text & ".00"

    End Sub

End Class
Imports System.Data.OleDb

Public Class frmStockSale
    Dim sSQL As String
    Dim Row As Integer ' Row Item Index


    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub frmStockSale_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate


    End Sub

    Private Sub frmStockSale_Leave(sender As Object, e As EventArgs) Handles Me.Leave

    End Sub

    Private Sub frmStockSale_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer

        For i = 0 To 9
            ClearControls(i)
        Next

        Row = 0
        CreateIDS(0)

    End Sub

    Private Sub CreateIDS(Index As Integer)
        Select Case Index
            Case 0
                CreateItemRow0()
            Case 1
                CreateItemRow1()
            Case 2
                CreateItemRow2()
            Case 3
                CreateItemRow3()
            Case 4
                CreateItemRow4()
            Case 5
                CreateItemRow5()
            Case 6
                CreateItemRow6()
            Case 7
                CreateItemRow7()
            Case 8
                CreateItemRow8()
            Case 9
                CreateItemRow9()
            Case Else
                MsgBox("Only 10 allowed")
        End Select

    End Sub

    Private Sub CreateItemRow0()
        Try

            cboPawnStockID0.Items.Clear()
            cboNormalStockID0.Items.Clear()

            'Populate comboboxes
            cboPawnStockID0.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID0.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            'cboPawnStockID0.SelectedText = "0"

            cboNormalStockID0.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID0.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            'cboNormalStockID0.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow1()
        Try

            cboNormalStockID1.Enabled = True
            cboPawnStockID1.Enabled = True
            cboQuantity1.Enabled = True
            txtDescription1.Enabled = True
            txtSerialNr1.Enabled = True
            txtPrice1.Enabled = True
            txtSubTotal1.Enabled = True
            bDelete1.Enabled = True
            bDelete1.Image = bDelete0.Image
            bSearch1.Enabled = True
            bSearch1.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID1.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID1.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID1.SelectedText = "0"

            cboNormalStockID1.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID1.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID1.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow2()
        Try

            cboNormalStockID2.Enabled = True
            cboPawnStockID2.Enabled = True
            cboQuantity2.Enabled = True
            txtDescription2.Enabled = True
            txtSerialNr2.Enabled = True
            txtPrice2.Enabled = True
            txtSubTotal2.Enabled = True
            bDelete2.Enabled = True
            bDelete2.Image = bDelete0.Image
            bSearch2.Enabled = True
            bSearch2.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID2.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID2.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID2.SelectedText = "0"

            cboNormalStockID2.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID2.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID2.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow3()
        Try
            cboNormalStockID3.Enabled = True
            cboPawnStockID3.Enabled = True
            cboQuantity3.Enabled = True
            txtDescription3.Enabled = True
            txtSerialNr3.Enabled = True
            txtPrice3.Enabled = True
            txtSubTotal3.Enabled = True
            bDelete3.Enabled = True
            bDelete3.Image = bDelete0.Image
            bSearch3.Enabled = True
            bSearch3.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID3.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID3.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID3.SelectedText = "0"

            cboNormalStockID3.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID3.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID3.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow4()
        Try
            cboNormalStockID4.Enabled = True
            cboPawnStockID4.Enabled = True
            cboQuantity4.Enabled = True
            txtDescription4.Enabled = True
            txtSerialNr4.Enabled = True
            txtPrice4.Enabled = True
            txtSubTotal4.Enabled = True
            bDelete4.Enabled = True
            bDelete4.Image = bDelete0.Image
            bSearch4.Enabled = True
            bSearch4.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID4.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID4.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID4.SelectedText = "0"

            cboNormalStockID4.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID4.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID4.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow5()
        Try
            cboNormalStockID5.Enabled = True
            cboPawnStockID5.Enabled = True
            cboQuantity5.Enabled = True
            txtDescription5.Enabled = True
            txtSerialNr5.Enabled = True
            txtPrice5.Enabled = True
            txtSubTotal5.Enabled = True
            bDelete5.Enabled = True
            bDelete5.Image = bDelete0.Image
            bSearch5.Enabled = True
            bSearch5.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID5.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID5.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID5.SelectedText = "0"

            cboNormalStockID5.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID5.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID5.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow6()
        Try
            cboNormalStockID6.Enabled = True
            cboPawnStockID6.Enabled = True
            cboQuantity6.Enabled = True
            txtDescription6.Enabled = True
            txtSerialNr6.Enabled = True
            txtPrice6.Enabled = True
            txtSubTotal6.Enabled = True
            bDelete6.Enabled = True
            bDelete6.Image = bDelete0.Image
            bSearch6.Enabled = True
            bSearch6.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID6.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID6.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID6.SelectedText = "0"

            cboNormalStockID6.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID6.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID6.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow7()
        Try
            cboNormalStockID7.Enabled = True
            cboPawnStockID7.Enabled = True
            cboQuantity7.Enabled = True
            txtDescription7.Enabled = True
            txtSerialNr7.Enabled = True
            txtPrice7.Enabled = True
            txtSubTotal7.Enabled = True
            bDelete7.Enabled = True
            bDelete7.Image = bDelete0.Image
            bSearch7.Enabled = True
            bSearch7.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID7.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID7.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID7.SelectedText = "0"

            cboNormalStockID7.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID7.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID7.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow8()
        Try
            cboNormalStockID8.Enabled = True
            cboPawnStockID8.Enabled = True
            cboQuantity8.Enabled = True
            txtDescription8.Enabled = True
            txtSerialNr8.Enabled = True
            txtPrice8.Enabled = True
            txtSubTotal8.Enabled = True
            bDelete8.Enabled = True
            bDelete8.Image = bDelete0.Image
            bSearch8.Enabled = True
            bSearch8.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID8.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID8.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID8.SelectedText = "0"

            cboNormalStockID8.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID8.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID8.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CreateItemRow9()
        Try
            cboNormalStockID9.Enabled = True
            cboPawnStockID9.Enabled = True
            cboQuantity9.Enabled = True
            txtDescription9.Enabled = True
            txtSerialNr9.Enabled = True
            txtPrice9.Enabled = True
            txtSubTotal9.Enabled = True
            bDelete9.Enabled = True
            bDelete9.Image = bDelete0.Image
            bSearch9.Enabled = True
            bSearch9.Image = bSearch0.Image

            'Populate comboboxes
            cboPawnStockID9.Items.Add("0")
            sSQL = "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                cboPawnStockID9.Items.Add(reader(0).ToString)
            End While
            reader.Close()
            ''Make first item (namely zero) selected
            cboPawnStockID9.SelectedText = "0"

            cboNormalStockID9.Items.Add("0")
            sSQL = "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                cboNormalStockID9.Items.Add(reader2(0).ToString)
            End While
            reader2.Close()
            ''Make first item (namely zero) selected
            cboNormalStockID9.SelectedText = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow0()
        Try

            'Clear comboboxes
            cboPawnStockID0.Text = "0"
            cboNormalStockID0.Text = "0"
            cboQuantity0.Items.Clear()
            cboQuantity0.Text = ""
            txtDescription0.Text = ""
            txtPrice0.Text = ""
            txtSerialNr0.Text = ""
            txtSubTotal0.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow1()
        Try
            cboNormalStockID1.Enabled = False
            cboPawnStockID1.Enabled = False
            cboQuantity1.Enabled = False
            txtDescription1.Enabled = False
            txtSerialNr1.Enabled = False
            txtPrice1.Enabled = False
            txtSubTotal1.Enabled = False
            bDelete1.Enabled = False
            bDelete1.Image = bDeleteDisable.Image
            bSearch1.Enabled = False
            bSearch1.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID1.Items.Clear()
            cboPawnStockID1.Text = ""
            cboNormalStockID1.Items.Clear()
            cboNormalStockID1.Text = ""

            cboQuantity1.Items.Clear()
            cboQuantity1.Text = ""
            txtDescription1.Text = ""
            txtPrice1.Text = ""
            txtSerialNr1.Text = ""
            txtSubTotal1.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow2()
        Try
            cboNormalStockID2.Enabled = False
            cboPawnStockID2.Enabled = False
            cboQuantity2.Enabled = False
            txtDescription2.Enabled = False
            txtSerialNr2.Enabled = False
            txtPrice2.Enabled = False
            txtSubTotal2.Enabled = False
            bDelete2.Enabled = False
            bDelete2.Image = bDeleteDisable.Image
            bSearch2.Enabled = False
            bSearch2.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID2.Items.Clear()
            cboPawnStockID2.Text = ""
            cboNormalStockID2.Items.Clear()
            cboNormalStockID2.Text = ""

            cboQuantity2.Items.Clear()
            cboQuantity2.Text = ""
            txtDescription2.Text = ""
            txtPrice2.Text = ""
            txtSerialNr2.Text = ""
            txtSubTotal2.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow3()
        Try
            cboNormalStockID3.Enabled = False
            cboPawnStockID3.Enabled = False
            cboQuantity3.Enabled = False
            txtDescription3.Enabled = False
            txtSerialNr3.Enabled = False
            txtPrice3.Enabled = False
            txtSubTotal3.Enabled = False
            bDelete3.Enabled = False
            bDelete3.Image = bDeleteDisable.Image
            bSearch3.Enabled = False
            bSearch3.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID3.Items.Clear()
            cboPawnStockID3.Text = ""
            cboNormalStockID3.Items.Clear()
            cboNormalStockID3.Text = ""

            cboQuantity3.Items.Clear()
            cboQuantity3.Text = ""
            txtDescription3.Text = ""
            txtPrice3.Text = ""
            txtSerialNr3.Text = ""
            txtSubTotal3.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow4()
        Try
            cboNormalStockID4.Enabled = False
            cboPawnStockID4.Enabled = False
            cboQuantity4.Enabled = False
            txtDescription4.Enabled = False
            txtSerialNr4.Enabled = False
            txtPrice4.Enabled = False
            txtSubTotal4.Enabled = False
            bDelete4.Enabled = False
            bDelete4.Image = bDeleteDisable.Image
            bSearch4.Enabled = False
            bSearch4.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID4.Items.Clear()
            cboPawnStockID4.Text = ""
            cboNormalStockID4.Items.Clear()
            cboNormalStockID4.Text = ""

            cboQuantity4.Items.Clear()
            cboQuantity4.Text = ""
            txtDescription4.Text = ""
            txtPrice4.Text = ""
            txtSerialNr4.Text = ""
            txtSubTotal4.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow5()
        Try
            cboNormalStockID5.Enabled = False
            cboPawnStockID5.Enabled = False
            cboQuantity5.Enabled = False
            txtDescription5.Enabled = False
            txtSerialNr5.Enabled = False
            txtPrice5.Enabled = False
            txtSubTotal5.Enabled = False
            bDelete5.Enabled = False
            bDelete5.Image = bDeleteDisable.Image
            bSearch5.Enabled = False
            bSearch5.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID5.Items.Clear()
            cboPawnStockID5.Text = ""
            cboNormalStockID5.Items.Clear()
            cboNormalStockID5.Text = ""

            cboQuantity5.Items.Clear()
            cboQuantity5.Text = ""
            txtDescription5.Text = ""
            txtPrice5.Text = ""
            txtSerialNr5.Text = ""
            txtSubTotal5.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow6()
        Try
            cboNormalStockID6.Enabled = False
            cboPawnStockID6.Enabled = False
            cboQuantity6.Enabled = False
            txtDescription6.Enabled = False
            txtSerialNr6.Enabled = False
            txtPrice6.Enabled = False
            txtSubTotal6.Enabled = False
            bDelete6.Enabled = False
            bDelete6.Image = bDeleteDisable.Image
            bSearch6.Enabled = False
            bSearch6.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID6.Items.Clear()
            cboPawnStockID6.Text = ""
            cboNormalStockID6.Items.Clear()
            cboNormalStockID6.Text = ""

            cboQuantity6.Items.Clear()
            cboQuantity6.Text = ""
            txtDescription6.Text = ""
            txtPrice6.Text = ""
            txtSerialNr6.Text = ""
            txtSubTotal6.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow7()
        Try
            cboNormalStockID7.Enabled = False
            cboPawnStockID7.Enabled = False
            cboQuantity7.Enabled = False
            txtDescription7.Enabled = False
            txtSerialNr7.Enabled = False
            txtPrice7.Enabled = False
            txtSubTotal7.Enabled = False
            bDelete7.Enabled = False
            bDelete7.Image = bDeleteDisable.Image
            bSearch7.Enabled = False
            bSearch7.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID7.Items.Clear()
            cboPawnStockID7.Text = ""
            cboNormalStockID7.Items.Clear()
            cboNormalStockID7.Text = ""

            cboQuantity7.Items.Clear()
            cboQuantity7.Text = ""
            txtDescription7.Text = ""
            txtPrice7.Text = ""
            txtSerialNr7.Text = ""
            txtSubTotal7.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow8()
        Try
            cboNormalStockID8.Enabled = False
            cboPawnStockID8.Enabled = False
            cboQuantity8.Enabled = False
            txtDescription8.Enabled = False
            txtSerialNr8.Enabled = False
            txtPrice8.Enabled = False
            txtSubTotal8.Enabled = False
            bDelete8.Enabled = False
            bDelete8.Image = bDeleteDisable.Image
            bSearch8.Enabled = False
            bSearch8.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID8.Items.Clear()
            cboPawnStockID8.Text = ""
            cboNormalStockID8.Items.Clear()
            cboNormalStockID8.Text = ""

            cboQuantity8.Items.Clear()
            cboQuantity8.Text = ""
            txtDescription8.Text = ""
            txtPrice8.Text = ""
            txtSerialNr8.Text = ""
            txtSubTotal8.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub RemoveItemRow9()
        Try
            cboNormalStockID9.Enabled = False
            cboPawnStockID9.Enabled = False
            cboQuantity9.Enabled = False
            txtDescription9.Enabled = False
            txtSerialNr9.Enabled = False
            txtPrice9.Enabled = False
            txtSubTotal9.Enabled = False
            bDelete9.Enabled = False
            bDelete9.Image = bDeleteDisable.Image
            bSearch9.Enabled = False
            bSearch9.Image = bSearchDisable.Image

            'Clear comboboxes
            cboPawnStockID9.Items.Clear()
            cboPawnStockID9.Text = ""
            cboNormalStockID9.Items.Clear()
            cboNormalStockID9.Text = ""

            cboQuantity9.Items.Clear()
            cboQuantity9.Text = ""
            txtDescription9.Text = ""
            txtPrice9.Text = ""
            txtSerialNr9.Text = ""
            txtSubTotal9.Text = ""

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub ClearControls(Index As Integer)
        ''Clear Controls
        Select Case Index
            Case 0
                RemoveItemRow0()
            Case 1
                RemoveItemRow1()
            Case 2
                RemoveItemRow2()
            Case 3
                RemoveItemRow3()
            Case 4
                RemoveItemRow4()
            Case 5
                RemoveItemRow5()
            Case 6
                RemoveItemRow6()
            Case 7
                RemoveItemRow7()
            Case 8
                RemoveItemRow8()
            Case 9
                RemoveItemRow9()
            Case Else

        End Select
    End Sub

    Private Sub cmdAddAnotherItem_Click(sender As Object, e As EventArgs) Handles cmdAddAnotherItem.Click

        Row = Row + 1
        If Row = 10 Then
            Row = Row - 1
            MsgBox("Only 10 transactions allowed.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If cboNormalStockID1.Enabled = False Then
            CreateIDS(1)
            Exit Sub
        End If

        If cboNormalStockID2.Enabled = False Then
            CreateIDS(2)
            Exit Sub
        End If

        If cboNormalStockID3.Enabled = False Then
            CreateIDS(3)
            Exit Sub
        End If

        If cboNormalStockID4.Enabled = False Then
            CreateIDS(4)
            Exit Sub
        End If

        If cboNormalStockID5.Enabled = False Then
            CreateIDS(5)
            Exit Sub
        End If

        If cboNormalStockID6.Enabled = False Then
            CreateIDS(6)
            Exit Sub
        End If

        If cboNormalStockID7.Enabled = False Then
            CreateIDS(7)
            Exit Sub
        End If

        If cboNormalStockID8.Enabled = False Then
            CreateIDS(8)
            Exit Sub
        End If

        If cboNormalStockID9.Enabled = False Then
            CreateIDS(9)
            Exit Sub
        End If

        'CreateIDS(row)

    End Sub

    Private Sub bDelete9_Click(sender As Object, e As EventArgs) Handles bDelete9.Click
        ClearControls(9)
        Row = Row - 1

    End Sub

    Private Sub bDelete8_Click(sender As Object, e As EventArgs) Handles bDelete8.Click
        ClearControls(8)
        Row = Row - 1

    End Sub

    Private Sub bDelete7_Click(sender As Object, e As EventArgs) Handles bDelete7.Click
        ClearControls(7)
        Row = Row - 1

    End Sub

    Private Sub bDelete6_Click(sender As Object, e As EventArgs) Handles bDelete6.Click
        ClearControls(6)
        Row = Row - 1

    End Sub

    Private Sub bDelete5_Click(sender As Object, e As EventArgs) Handles bDelete5.Click
        ClearControls(5)
        Row = Row - 1

    End Sub

    Private Sub bDelete4_Click(sender As Object, e As EventArgs) Handles bDelete4.Click
        ClearControls(4)
        Row = Row - 1

    End Sub

    Private Sub bDelete3_Click(sender As Object, e As EventArgs) Handles bDelete3.Click
        ClearControls(3)
        Row = Row - 1

    End Sub

    Private Sub bDelete2_Click(sender As Object, e As EventArgs) Handles bDelete2.Click
        ClearControls(2)
        Row = Row - 1

    End Sub

    Private Sub bDelete1_Click(sender As Object, e As EventArgs) Handles bDelete1.Click
        ClearControls(1)
        Row = Row - 1

    End Sub

    Private Sub bDelete0_Click(sender As Object, e As EventArgs) Handles bDelete0.Click
        ClearControls(0)

    End Sub

    Private Sub cboPawnStockID0_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID0.SelectedIndexChanged
        Try

            If cboPawnStockID0.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID0.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity0.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity0.Items.Add(i)
            Next i

            cboQuantity0.SelectedIndex = 0

            'Populate other controls
            txtDescription0.Text = reader(1).ToString
            txtSerialNr0.Text = reader(2).ToString
            txtPrice0.Text = reader(3).ToString
            SetCurr(txtPrice0)
            txtPrice0.Enabled = True
            reader.Close()

            cboNormalStockID0.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity0_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity0.SelectedIndexChanged
        If cboQuantity0.Text = "" Or txtPrice0.Text = "" Then Exit Sub

        txtSubTotal0.Text = cboQuantity0.Text * txtPrice0.Text
        SetCurr(txtSubTotal0)

    End Sub

    Private Sub txtPrice0_LostFocus(sender As Object, e As EventArgs) Handles txtPrice0.LostFocus
        SetCurr(txtPrice0)
    End Sub

    Private Sub txtPrice0_TextChanged(sender As Object, e As EventArgs) Handles txtPrice0.TextChanged
        If cboQuantity0.Text = "" Or txtPrice0.Text = "" Then Exit Sub

        txtSubTotal0.Text = cboQuantity0.Text * txtPrice0.Text
        SetCurr(txtSubTotal0)

    End Sub

    Private Sub cboNormalStockID0_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID0.SelectedIndexChanged
        Try

            If cboNormalStockID0.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID0.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity0.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity0.Items.Add(i)
            Next i

            cboQuantity0.SelectedIndex = 0

            'Populate other controls
            txtDescription0.Text = reader(1).ToString
            txtSerialNr0.Text = ""
            txtPrice0.Text = reader(2).ToString
            SetCurr(txtPrice0)
            txtPrice0.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories0.Checked = False
            Else
                chkCellphoneAssessories0.Checked = True
            End If

            reader.Close()

            cboPawnStockID0.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboPawnStockID1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID1.SelectedIndexChanged
        Try

            If cboPawnStockID1.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID1.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity1.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity1.Items.Add(i)
            Next i

            cboQuantity1.SelectedIndex = 0

            'Populate other controls
            txtDescription1.Text = reader(1).ToString
            txtSerialNr1.Text = reader(2).ToString
            txtPrice1.Text = reader(3).ToString
            SetCurr(txtPrice1)
            txtPrice1.Enabled = True
            reader.Close()

            cboNormalStockID1.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity1.SelectedIndexChanged
        If cboQuantity1.Text = "" Or txtPrice1.Text = "" Then Exit Sub

        txtSubTotal1.Text = cboQuantity1.Text * txtPrice1.Text
        SetCurr(txtSubTotal1)

    End Sub

    Private Sub txtPrice1_LostFocus(sender As Object, e As EventArgs) Handles txtPrice1.LostFocus
        SetCurr(txtPrice1)
    End Sub

    Private Sub txtPrice1_TextChanged(sender As Object, e As EventArgs) Handles txtPrice1.TextChanged
        If cboQuantity1.Text = "" Or txtPrice1.Text = "" Then Exit Sub

        txtSubTotal1.Text = cboQuantity1.Text * txtPrice1.Text
        SetCurr(txtSubTotal1)

    End Sub

    Private Sub cboNormalStockID1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID1.SelectedIndexChanged
        Try

            If cboNormalStockID1.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID1.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity1.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity1.Items.Add(i)
            Next i

            cboQuantity1.SelectedIndex = 0

            'Populate other controls
            txtDescription1.Text = reader(1).ToString
            txtSerialNr1.Text = ""
            txtPrice1.Text = reader(2).ToString
            SetCurr(txtPrice1)
            txtPrice1.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories1.Checked = False
            Else
                chkCellphoneAssessories1.Checked = True
            End If

            reader.Close()

            cboPawnStockID1.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal0_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal0.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub txtSubTotal1_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal1.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID2.SelectedIndexChanged
        Try

            If cboPawnStockID2.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID2.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity2.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity2.Items.Add(i)
            Next i

            cboQuantity2.SelectedIndex = 0

            'Populate other controls
            txtDescription2.Text = reader(1).ToString
            txtSerialNr2.Text = reader(2).ToString
            txtPrice2.Text = reader(3).ToString
            SetCurr(txtPrice2)
            txtPrice2.Enabled = True
            reader.Close()

            cboNormalStockID2.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity2.SelectedIndexChanged
        If cboQuantity2.Text = "" Or txtPrice2.Text = "" Then Exit Sub

        txtSubTotal2.Text = cboQuantity2.Text * txtPrice2.Text
        SetCurr(txtSubTotal2)

    End Sub

    Private Sub txtPrice2_LostFocus(sender As Object, e As EventArgs) Handles txtPrice2.LostFocus
        SetCurr(txtPrice2)
    End Sub

    Private Sub txtPrice2_TextChanged(sender As Object, e As EventArgs) Handles txtPrice2.TextChanged
        If cboQuantity2.Text = "" Or txtPrice2.Text = "" Then Exit Sub

        txtSubTotal2.Text = cboQuantity2.Text * txtPrice2.Text
        SetCurr(txtSubTotal2)

    End Sub

    Private Sub cboNormalStockID2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID2.SelectedIndexChanged
        Try

            If cboNormalStockID2.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID2.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity2.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity2.Items.Add(i)
            Next i

            cboQuantity2.SelectedIndex = 0

            'Populate other controls
            txtDescription2.Text = reader(1).ToString
            txtSerialNr2.Text = ""
            txtPrice2.Text = reader(2).ToString
            SetCurr(txtPrice2)
            txtPrice2.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories2.Checked = False
            Else
                chkCellphoneAssessories2.Checked = True
            End If

            reader.Close()

            cboPawnStockID2.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal2_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal2.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID3.SelectedIndexChanged
        Try

            If cboPawnStockID3.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID3.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity3.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity3.Items.Add(i)
            Next i

            cboQuantity3.SelectedIndex = 0

            'Populate other controls
            txtDescription3.Text = reader(1).ToString
            txtSerialNr3.Text = reader(2).ToString
            txtPrice3.Text = reader(3).ToString
            SetCurr(txtPrice3)
            txtPrice3.Enabled = True
            reader.Close()

            cboNormalStockID3.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity3.SelectedIndexChanged
        If cboQuantity3.Text = "" Or txtPrice3.Text = "" Then Exit Sub

        txtSubTotal3.Text = cboQuantity3.Text * txtPrice3.Text
        SetCurr(txtSubTotal3)

    End Sub

    Private Sub txtPrice3_LostFocus(sender As Object, e As EventArgs) Handles txtPrice3.LostFocus
        SetCurr(txtPrice3)
    End Sub

    Private Sub txtPrice3_TextChanged(sender As Object, e As EventArgs) Handles txtPrice3.TextChanged
        If cboQuantity3.Text = "" Or txtPrice3.Text = "" Then Exit Sub

        txtSubTotal3.Text = cboQuantity3.Text * txtPrice3.Text
        SetCurr(txtSubTotal3)

    End Sub

    Private Sub cboNormalStockID3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID3.SelectedIndexChanged
        Try

            If cboNormalStockID3.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID3.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity3.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity3.Items.Add(i)
            Next i

            cboQuantity3.SelectedIndex = 0

            'Populate other controls
            txtDescription3.Text = reader(1).ToString
            txtSerialNr3.Text = ""
            txtPrice3.Text = reader(2).ToString
            SetCurr(txtPrice3)
            txtPrice3.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories3.Checked = False
            Else
                chkCellphoneAssessories3.Checked = True
            End If

            reader.Close()

            cboPawnStockID3.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal3_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal3.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID4.SelectedIndexChanged
        Try

            If cboPawnStockID4.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID4.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity4.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity4.Items.Add(i)
            Next i

            cboQuantity4.SelectedIndex = 0

            'Populate other controls
            txtDescription4.Text = reader(1).ToString
            txtSerialNr4.Text = reader(2).ToString
            txtPrice4.Text = reader(3).ToString
            SetCurr(txtPrice4)
            txtPrice4.Enabled = True
            reader.Close()

            cboNormalStockID4.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity4.SelectedIndexChanged
        If cboQuantity4.Text = "" Or txtPrice4.Text = "" Then Exit Sub

        txtSubTotal4.Text = cboQuantity4.Text * txtPrice4.Text
        SetCurr(txtSubTotal4)

    End Sub

    Private Sub txtPrice4_LostFocus(sender As Object, e As EventArgs) Handles txtPrice4.LostFocus
        SetCurr(txtPrice4)
    End Sub

    Private Sub txtPrice4_TextChanged(sender As Object, e As EventArgs) Handles txtPrice4.TextChanged
        If cboQuantity4.Text = "" Or txtPrice4.Text = "" Then Exit Sub

        txtSubTotal4.Text = cboQuantity4.Text * txtPrice4.Text
        SetCurr(txtSubTotal4)

    End Sub

    Private Sub cboNormalStockID4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID4.SelectedIndexChanged
        Try

            If cboNormalStockID4.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID4.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity4.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity4.Items.Add(i)
            Next i

            cboQuantity4.SelectedIndex = 0

            'Populate other controls
            txtDescription4.Text = reader(1).ToString
            txtSerialNr4.Text = ""
            txtPrice4.Text = reader(2).ToString
            SetCurr(txtPrice4)
            txtPrice4.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories4.Checked = False
            Else
                chkCellphoneAssessories4.Checked = True
            End If

            reader.Close()

            cboPawnStockID4.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal4_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal4.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID5.SelectedIndexChanged
        Try

            If cboPawnStockID5.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID5.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity5.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity5.Items.Add(i)
            Next i

            cboQuantity5.SelectedIndex = 0

            'Populate other controls
            txtDescription5.Text = reader(1).ToString
            txtSerialNr5.Text = reader(2).ToString
            txtPrice5.Text = reader(3).ToString
            SetCurr(txtPrice5)
            txtPrice5.Enabled = True
            reader.Close()

            cboNormalStockID5.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity5.SelectedIndexChanged
        If cboQuantity5.Text = "" Or txtPrice5.Text = "" Then Exit Sub

        txtSubTotal5.Text = cboQuantity5.Text * txtPrice5.Text
        SetCurr(txtSubTotal5)

    End Sub

    Private Sub txtPrice5_LostFocus(sender As Object, e As EventArgs) Handles txtPrice5.LostFocus
        SetCurr(txtPrice5)
    End Sub

    Private Sub txtPrice5_TextChanged(sender As Object, e As EventArgs) Handles txtPrice5.TextChanged
        If cboQuantity5.Text = "" Or txtPrice5.Text = "" Then Exit Sub

        txtSubTotal5.Text = cboQuantity5.Text * txtPrice5.Text
        SetCurr(txtSubTotal5)

    End Sub

    Private Sub cboNormalStockID5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID5.SelectedIndexChanged
        Try

            If cboNormalStockID5.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID5.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity5.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity5.Items.Add(i)
            Next i

            cboQuantity5.SelectedIndex = 0

            'Populate other controls
            txtDescription5.Text = reader(1).ToString
            txtSerialNr5.Text = ""
            txtPrice5.Text = reader(2).ToString
            SetCurr(txtPrice5)
            txtPrice5.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories5.Checked = False
            Else
                chkCellphoneAssessories5.Checked = True
            End If

            reader.Close()

            cboPawnStockID5.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal5_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal5.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID6.SelectedIndexChanged
        Try

            If cboPawnStockID6.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID6.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity6.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity6.Items.Add(i)
            Next i

            cboQuantity6.SelectedIndex = 0

            'Populate other controls
            txtDescription6.Text = reader(1).ToString
            txtSerialNr6.Text = reader(2).ToString
            txtPrice6.Text = reader(3).ToString
            SetCurr(txtPrice6)
            txtPrice6.Enabled = True
            reader.Close()

            cboNormalStockID6.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity6.SelectedIndexChanged
        If cboQuantity6.Text = "" Or txtPrice6.Text = "" Then Exit Sub

        txtSubTotal6.Text = cboQuantity6.Text * txtPrice6.Text
        SetCurr(txtSubTotal6)

    End Sub

    Private Sub txtPrice6_LostFocus(sender As Object, e As EventArgs) Handles txtPrice6.LostFocus
        SetCurr(txtPrice6)
    End Sub

    Private Sub txtPrice6_TextChanged(sender As Object, e As EventArgs) Handles txtPrice6.TextChanged
        If cboQuantity6.Text = "" Or txtPrice6.Text = "" Then Exit Sub

        txtSubTotal6.Text = cboQuantity6.Text * txtPrice6.Text
        SetCurr(txtSubTotal6)

    End Sub

    Private Sub cboNormalStockID6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID6.SelectedIndexChanged
        Try

            If cboNormalStockID6.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID6.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity6.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity6.Items.Add(i)
            Next i

            cboQuantity6.SelectedIndex = 0

            'Populate other controls
            txtDescription6.Text = reader(1).ToString
            txtSerialNr6.Text = ""
            txtPrice6.Text = reader(2).ToString
            SetCurr(txtPrice6)
            txtPrice6.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories6.Checked = False
            Else
                chkCellphoneAssessories6.Checked = True
            End If

            reader.Close()

            cboPawnStockID6.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal6_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal6.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID7.SelectedIndexChanged
        Try

            If cboPawnStockID7.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID7.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity7.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity7.Items.Add(i)
            Next i

            cboQuantity7.SelectedIndex = 0

            'Populate other controls
            txtDescription7.Text = reader(1).ToString
            txtSerialNr7.Text = reader(2).ToString
            txtPrice7.Text = reader(3).ToString
            SetCurr(txtPrice7)
            txtPrice7.Enabled = True
            reader.Close()

            cboNormalStockID7.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity7.SelectedIndexChanged
        If cboQuantity7.Text = "" Or txtPrice7.Text = "" Then Exit Sub

        txtSubTotal7.Text = cboQuantity7.Text * txtPrice7.Text
        SetCurr(txtSubTotal7)

    End Sub

    Private Sub txtPrice7_LostFocus(sender As Object, e As EventArgs) Handles txtPrice7.LostFocus
        SetCurr(txtPrice7)
    End Sub

    Private Sub txtPrice7_TextChanged(sender As Object, e As EventArgs) Handles txtPrice7.TextChanged
        If cboQuantity7.Text = "" Or txtPrice7.Text = "" Then Exit Sub

        txtSubTotal7.Text = cboQuantity7.Text * txtPrice7.Text
        SetCurr(txtSubTotal7)

    End Sub

    Private Sub cboNormalStockID7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID7.SelectedIndexChanged
        Try

            If cboNormalStockID7.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID7.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity7.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity7.Items.Add(i)
            Next i

            cboQuantity7.SelectedIndex = 0

            'Populate other controls
            txtDescription7.Text = reader(1).ToString
            txtSerialNr7.Text = ""
            txtPrice7.Text = reader(2).ToString
            SetCurr(txtPrice7)
            txtPrice7.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories7.Checked = False
            Else
                chkCellphoneAssessories7.Checked = True
            End If

            reader.Close()

            cboPawnStockID7.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal7_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal7.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID8.SelectedIndexChanged
        Try

            If cboPawnStockID8.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID8.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity8.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity8.Items.Add(i)
            Next i

            cboQuantity8.SelectedIndex = 0

            'Populate other controls
            txtDescription8.Text = reader(1).ToString
            txtSerialNr8.Text = reader(2).ToString
            txtPrice8.Text = reader(3).ToString
            SetCurr(txtPrice8)
            txtPrice8.Enabled = True
            reader.Close()

            cboNormalStockID8.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity8.SelectedIndexChanged
        If cboQuantity8.Text = "" Or txtPrice8.Text = "" Then Exit Sub

        txtSubTotal8.Text = cboQuantity8.Text * txtPrice8.Text
        SetCurr(txtSubTotal8)

    End Sub

    Private Sub txtPrice8_LostFocus(sender As Object, e As EventArgs) Handles txtPrice8.LostFocus
        SetCurr(txtPrice8)
    End Sub

    Private Sub txtPrice8_TextChanged(sender As Object, e As EventArgs) Handles txtPrice8.TextChanged
        If cboQuantity8.Text = "" Or txtPrice8.Text = "" Then Exit Sub

        txtSubTotal8.Text = cboQuantity8.Text * txtPrice8.Text
        SetCurr(txtSubTotal8)

    End Sub

    Private Sub cboNormalStockID8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID8.SelectedIndexChanged
        Try

            If cboNormalStockID8.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID8.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity8.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity8.Items.Add(i)
            Next i

            cboQuantity8.SelectedIndex = 0

            'Populate other controls
            txtDescription8.Text = reader(1).ToString
            txtSerialNr8.Text = ""
            txtPrice8.Text = reader(2).ToString
            SetCurr(txtPrice8)
            txtPrice8.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories8.Checked = False
            Else
                chkCellphoneAssessories8.Checked = True
            End If

            reader.Close()

            cboPawnStockID8.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal8_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal8.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub cboPawnStockID9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPawnStockID9.SelectedIndexChanged
        Try

            If cboPawnStockID9.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID9.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity9.Items.Clear()

            Dim i As Integer
            reader.Read()
            For i = 1 To reader(0).ToString
                cboQuantity9.Items.Add(i)
            Next i

            cboQuantity9.SelectedIndex = 0

            'Populate other controls
            txtDescription9.Text = reader(1).ToString
            txtSerialNr9.Text = reader(2).ToString
            txtPrice9.Text = reader(3).ToString
            SetCurr(txtPrice9)
            txtPrice9.Enabled = True
            reader.Close()

            cboNormalStockID9.Text = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cboQuantity9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboQuantity9.SelectedIndexChanged
        If cboQuantity9.Text = "" Or txtPrice9.Text = "" Then Exit Sub

        txtSubTotal9.Text = cboQuantity9.Text * txtPrice9.Text
        SetCurr(txtSubTotal9)

    End Sub

    Private Sub txtPrice9_LostFocus(sender As Object, e As EventArgs) Handles txtPrice9.LostFocus
        SetCurr(txtPrice9)
    End Sub

    Private Sub txtPrice9_TextChanged(sender As Object, e As EventArgs) Handles txtPrice9.TextChanged
        If cboQuantity9.Text = "" Or txtPrice9.Text = "" Then Exit Sub

        txtSubTotal9.Text = cboQuantity9.Text * txtPrice9.Text
        SetCurr(txtSubTotal9)

    End Sub

    Private Sub cboNormalStockID9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNormalStockID9.SelectedIndexChanged
        Try

            If cboNormalStockID9.Text = 0 Then Exit Sub

            'Retrieve relevant information and populate controls
            sSQL = "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID9.Text
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()

            'Empty quantity combobox and then repopulate
            cboQuantity9.Items.Clear()
            Dim i As Integer
            reader.Read()


            For i = 1 To reader(0).ToString
                cboQuantity9.Items.Add(i)
            Next i

            cboQuantity9.SelectedIndex = 0

            'Populate other controls
            txtDescription9.Text = reader(1).ToString
            txtSerialNr9.Text = ""
            txtPrice9.Text = reader(2).ToString
            SetCurr(txtPrice9)
            txtPrice9.Enabled = True

            'Populate other controls
            If reader(3).ToString = "0" Then
                chkCellphoneAssessories9.Checked = False
            Else
                chkCellphoneAssessories9.Checked = True
            End If

            reader.Close()

            cboPawnStockID9.SelectedIndex = "0"

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub txtSubTotal9_TextChanged(sender As Object, e As EventArgs) Handles txtSubTotal9.TextChanged
        Dim iInt As Integer

        iInt = 0
        If txtSubTotal0.Text <> "" Then iInt = iInt + CInt(txtSubTotal0.Text)
        If txtSubTotal1.Text <> "" Then iInt = iInt + CInt(txtSubTotal1.Text)
        If txtSubTotal2.Text <> "" Then iInt = iInt + CInt(txtSubTotal2.Text)
        If txtSubTotal3.Text <> "" Then iInt = iInt + CInt(txtSubTotal3.Text)
        If txtSubTotal4.Text <> "" Then iInt = iInt + CInt(txtSubTotal4.Text)
        If txtSubTotal5.Text <> "" Then iInt = iInt + CInt(txtSubTotal5.Text)
        If txtSubTotal6.Text <> "" Then iInt = iInt + CInt(txtSubTotal6.Text)
        If txtSubTotal7.Text <> "" Then iInt = iInt + CInt(txtSubTotal7.Text)
        If txtSubTotal8.Text <> "" Then iInt = iInt + CInt(txtSubTotal8.Text)
        If txtSubTotal9.Text <> "" Then iInt = iInt + CInt(txtSubTotal9.Text)

        txtGrandTotal.Text = iInt
        SetCurr(txtGrandTotal)


    End Sub

    Private Sub AddSearchItem0()
        If p_Search_NSID <> "" Then
            cboNormalStockID0.Text = p_Search_NSID
            cboNormalStockID0_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID0.Text = p_Search_PSID
            cboPawnStockID0_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID0.Text = "0"
            cboNormalStockID0.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem1()
        If p_Search_NSID <> "" Then
            cboNormalStockID1.Text = p_Search_NSID
            cboNormalStockID1_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID1.Text = p_Search_PSID
            cboPawnStockID1_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID1.Text = "0"
            cboNormalStockID1.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem2()
        If p_Search_NSID <> "" Then
            cboNormalStockID2.Text = p_Search_NSID
            cboNormalStockID2_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID2.Text = p_Search_PSID
            cboPawnStockID2_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID2.Text = "0"
            cboNormalStockID2.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem3()
        If p_Search_NSID <> "" Then
            cboNormalStockID3.Text = p_Search_NSID
            cboNormalStockID3_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID3.Text = p_Search_PSID
            cboPawnStockID3_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID3.Text = "0"
            cboNormalStockID3.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem4()
        If p_Search_NSID <> "" Then
            cboNormalStockID4.Text = p_Search_NSID
            cboNormalStockID4_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID4.Text = p_Search_PSID
            cboPawnStockID4_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID4.Text = "0"
            cboNormalStockID4.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem5()
        If p_Search_NSID <> "" Then
            cboNormalStockID5.Text = p_Search_NSID
            cboNormalStockID5_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID5.Text = p_Search_PSID
            cboPawnStockID5_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID5.Text = "0"
            cboNormalStockID5.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem6()
        If p_Search_NSID <> "" Then
            cboNormalStockID6.Text = p_Search_NSID
            cboNormalStockID6_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID6.Text = p_Search_PSID
            cboPawnStockID6_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID6.Text = "0"
            cboNormalStockID6.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem7()
        If p_Search_NSID <> "" Then
            cboNormalStockID7.Text = p_Search_NSID
            cboNormalStockID7_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID7.Text = p_Search_PSID
            cboPawnStockID7_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID7.Text = "0"
            cboNormalStockID7.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem8()
        If p_Search_NSID <> "" Then
            cboNormalStockID8.Text = p_Search_NSID
            cboNormalStockID8_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID8.Text = p_Search_PSID
            cboPawnStockID8_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID8.Text = "0"
            cboNormalStockID8.Text = "0"
        End If

    End Sub

    Private Sub AddSearchItem9()
        If p_Search_NSID <> "" Then
            cboNormalStockID9.Text = p_Search_NSID
            cboNormalStockID9_SelectedIndexChanged(Nothing, Nothing)
        ElseIf p_Search_PSID <> "" Then
            cboPawnStockID9.Text = p_Search_PSID
            cboPawnStockID9_SelectedIndexChanged(Nothing, Nothing)
        Else
            cboPawnStockID9.Text = "0"
            cboNormalStockID9.Text = "0"
        End If

    End Sub

    Private Sub bSearch0_Click(sender As Object, e As EventArgs) Handles bSearch0.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem0()
        End If

    End Sub

    Private Sub bSearch1_Click(sender As Object, e As EventArgs) Handles bSearch1.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem1()
        End If

    End Sub

    Private Sub bSearch2_Click(sender As Object, e As EventArgs) Handles bSearch2.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem2()
        End If

    End Sub

    Private Sub bSearch3_Click(sender As Object, e As EventArgs) Handles bSearch3.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem3()
        End If

    End Sub

    Private Sub bSearch4_Click(sender As Object, e As EventArgs) Handles bSearch4.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem4()
        End If

    End Sub

    Private Sub bSearch5_Click(sender As Object, e As EventArgs) Handles bSearch5.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem5()
        End If

    End Sub

    Private Sub bSearch6_Click(sender As Object, e As EventArgs) Handles bSearch6.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem6()
        End If

    End Sub

    Private Sub bSearch7_Click(sender As Object, e As EventArgs) Handles bSearch7.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem7()
        End If

    End Sub

    Private Sub bSearch8_Click(sender As Object, e As EventArgs) Handles bSearch8.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem8()
        End If

    End Sub

    Private Sub bSearch9_Click(sender As Object, e As EventArgs) Handles bSearch9.Click
        frmStockSearch.ShowDialog(Me)
        If p_Search_NSID <> "" Or p_Search_PSID <> "" Then
            AddSearchItem9()
        End If

    End Sub

    Public Sub AddItemsArray(ItemCount As Integer, sItemId As String, sQuatity As String, sDescription As String, sSerialnr As String, sPrice As String, sSubTotal As String, sSelDate As String)

        If sPrice <> "" Then sPrice = p_Currency & sPrice
        If sSubTotal <> "" Then sSubTotal = p_Currency & sSubTotal

        arrItems(ItemCount, 0) = sItemId
        arrItems(ItemCount, 1) = sQuatity
        arrItems(ItemCount, 2) = sDescription
        arrItems(ItemCount, 3) = sSerialnr
        arrItems(ItemCount, 4) = sPrice
        arrItems(ItemCount, 5) = sSubTotal
        arrItems(ItemCount, 6) = sSelDate

    End Sub

    Private Sub cmdTransact_Click(sender As Object, e As EventArgs) Handles cmdTransact.Click
        Dim i As Integer
        Dim bAtLeastOne As Boolean

        bAtLeastOne = False
        iSaleItemCount = 0

        For i = 0 To 9
            AddItemsArray(i, "", "", "", "", "", "", "")
        Next


        For i = 0 To 9
            Select Case i
                Case 0
                    If cboNormalStockID0.Text <> "0" Then
                        NormalTransAction(cboNormalStockID0.Text, cboQuantity0.Text, txtDescription0.Text, txtSubTotal0.Text, txtPrice0.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories0.Checked)
                        bAtLeastOne = True
                        AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID0.Text, cboQuantity0.Text, txtDescription0.Text, txtSerialNr0.Text, txtPrice0.Text, txtSubTotal0.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                        iSaleItemCount = iSaleItemCount + 1
                    ElseIf cboPawnStockID0.Text <> "0" Then
                        PawnTransAction(cboPawnStockID0.Text, cboQuantity0.Text, txtDescription0.Text, txtSubTotal0.Text, txtPrice0.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                        bAtLeastOne = True
                        AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID0.Text, cboQuantity0.Text, txtDescription0.Text, txtSerialNr0.Text, txtPrice0.Text, txtSubTotal0.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                        iSaleItemCount = iSaleItemCount + 1
                    End If
                Case 1
                    If cboNormalStockID1.Enabled = True Then
                        If cboNormalStockID1.Text <> "0" Then
                            NormalTransAction(cboNormalStockID1.Text, cboQuantity1.Text, txtDescription1.Text, txtSubTotal1.Text, txtPrice1.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories1.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID1.Text, cboQuantity1.Text, txtDescription1.Text, txtSerialNr1.Text, txtPrice1.Text, txtSubTotal1.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID1.Text <> "0" Then
                            PawnTransAction(cboPawnStockID1.Text, cboQuantity1.Text, txtDescription1.Text, txtSubTotal1.Text, txtPrice1.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID1.Text, cboQuantity1.Text, txtDescription1.Text, txtSerialNr1.Text, txtPrice1.Text, txtSubTotal1.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 2
                    If cboNormalStockID2.Enabled = True Then
                        If cboNormalStockID2.Text <> "0" Then
                            NormalTransAction(cboNormalStockID2.Text, cboQuantity2.Text, txtDescription2.Text, txtSubTotal2.Text, txtPrice2.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories2.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID2.Text, cboQuantity2.Text, txtDescription2.Text, txtSerialNr2.Text, txtPrice2.Text, txtSubTotal2.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID2.Text <> "0" Then
                            PawnTransAction(cboPawnStockID2.Text, cboQuantity2.Text, txtDescription2.Text, txtSubTotal2.Text, txtPrice2.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID2.Text, cboQuantity2.Text, txtDescription2.Text, txtSerialNr2.Text, txtPrice2.Text, txtSubTotal2.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 3
                    If cboNormalStockID3.Enabled = True Then
                        If cboNormalStockID3.Text <> "0" Then
                            NormalTransAction(cboNormalStockID3.Text, cboQuantity3.Text, txtDescription3.Text, txtSubTotal3.Text, txtPrice3.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories3.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID3.Text, cboQuantity3.Text, txtDescription3.Text, txtSerialNr3.Text, txtPrice3.Text, txtSubTotal3.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID3.Text <> "0" Then
                            PawnTransAction(cboPawnStockID3.Text, cboQuantity3.Text, txtDescription3.Text, txtSubTotal3.Text, txtPrice3.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID3.Text, cboQuantity3.Text, txtDescription3.Text, txtSerialNr3.Text, txtPrice3.Text, txtSubTotal3.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 4
                    If cboNormalStockID4.Enabled = True Then
                        If cboNormalStockID4.Text <> "0" Then
                            NormalTransAction(cboNormalStockID4.Text, cboQuantity4.Text, txtDescription4.Text, txtSubTotal4.Text, txtPrice4.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories4.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID4.Text, cboQuantity4.Text, txtDescription4.Text, txtSerialNr4.Text, txtPrice4.Text, txtSubTotal4.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID4.Text <> "0" Then
                            PawnTransAction(cboPawnStockID4.Text, cboQuantity4.Text, txtDescription4.Text, txtSubTotal4.Text, txtPrice4.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID4.Text, cboQuantity4.Text, txtDescription4.Text, txtSerialNr4.Text, txtPrice4.Text, txtSubTotal4.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 5
                    If cboNormalStockID5.Enabled = True Then
                        If cboNormalStockID5.Text <> "0" Then
                            NormalTransAction(cboNormalStockID5.Text, cboQuantity5.Text, txtDescription5.Text, txtSubTotal5.Text, txtPrice5.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories5.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID5.Text, cboQuantity5.Text, txtDescription5.Text, txtSerialNr5.Text, txtPrice5.Text, txtSubTotal5.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID5.Text <> "0" Then
                            PawnTransAction(cboPawnStockID5.Text, cboQuantity5.Text, txtDescription5.Text, txtSubTotal5.Text, txtPrice5.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID5.Text, cboQuantity5.Text, txtDescription5.Text, txtSerialNr5.Text, txtPrice5.Text, txtSubTotal5.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 6
                    If cboNormalStockID6.Enabled = True Then
                        If cboNormalStockID6.Text <> "0" Then
                            NormalTransAction(cboNormalStockID6.Text, cboQuantity6.Text, txtDescription6.Text, txtSubTotal6.Text, txtPrice6.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories6.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID6.Text, cboQuantity6.Text, txtDescription6.Text, txtSerialNr6.Text, txtPrice6.Text, txtSubTotal6.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID6.Text <> "0" Then
                            PawnTransAction(cboPawnStockID6.Text, cboQuantity6.Text, txtDescription6.Text, txtSubTotal6.Text, txtPrice6.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID6.Text, cboQuantity6.Text, txtDescription6.Text, txtSerialNr6.Text, txtPrice6.Text, txtSubTotal6.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 7
                    If cboNormalStockID7.Enabled = True Then
                        If cboNormalStockID7.Text <> "0" Then
                            NormalTransAction(cboNormalStockID7.Text, cboQuantity7.Text, txtDescription7.Text, txtSubTotal7.Text, txtPrice7.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories7.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID7.Text, cboQuantity7.Text, txtDescription7.Text, txtSerialNr7.Text, txtPrice7.Text, txtSubTotal7.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID7.Text <> "0" Then
                            PawnTransAction(cboPawnStockID7.Text, cboQuantity7.Text, txtDescription7.Text, txtSubTotal7.Text, txtPrice7.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID7.Text, cboQuantity7.Text, txtDescription7.Text, txtSerialNr7.Text, txtPrice7.Text, txtSubTotal7.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 8
                    If cboNormalStockID8.Enabled = True Then
                        If cboNormalStockID8.Text <> "0" Then
                            NormalTransAction(cboNormalStockID8.Text, cboQuantity8.Text, txtDescription8.Text, txtSubTotal8.Text, txtPrice8.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories8.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID8.Text, cboQuantity8.Text, txtDescription8.Text, txtSerialNr8.Text, txtPrice8.Text, txtSubTotal8.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID8.Text <> "0" Then
                            PawnTransAction(cboPawnStockID8.Text, cboQuantity8.Text, txtDescription8.Text, txtSubTotal8.Text, txtPrice8.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID8.Text, cboQuantity8.Text, txtDescription8.Text, txtSerialNr8.Text, txtPrice8.Text, txtSubTotal8.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case 9
                    If cboNormalStockID9.Enabled = True Then
                        If cboNormalStockID9.Text <> "0" Then
                            NormalTransAction(cboNormalStockID9.Text, cboQuantity9.Text, txtDescription9.Text, txtSubTotal9.Text, txtPrice9.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate), chkCellphoneAssessories9.Checked)
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "NS_ID: " & cboNormalStockID9.Text, cboQuantity9.Text, txtDescription9.Text, txtSerialNr9.Text, txtPrice9.Text, txtSubTotal9.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        ElseIf cboPawnStockID9.Text <> "0" Then
                            PawnTransAction(cboPawnStockID9.Text, cboQuantity9.Text, txtDescription9.Text, txtSubTotal9.Text, txtPrice9.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            bAtLeastOne = True
                            AddItemsArray(iSaleItemCount, "PS_ID: " & cboPawnStockID9.Text, cboQuantity9.Text, txtDescription9.Text, txtSerialNr9.Text, txtPrice9.Text, txtSubTotal9.Text, FormatDateTime(dtTranDate.Value, DateFormat.ShortDate))
                            iSaleItemCount = iSaleItemCount + 1
                        End If
                    End If
                Case Else
                    'do nothing
            End Select

        Next

        If bAtLeastOne = False Then
            MsgBox("No stock selected for transaction.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Exit Sub
        Else
            sSaleTotal = p_Currency & txtGrandTotal.Text
            With frmReports
                .Label1.Text = "Generating Stock Sale Invoice ..."
                .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
                .Height = 67
                .Report = frmReports.ReportType.StockSale
                .ShowDialog(Me)
            End With
            Me.Close()
        End If


    End Sub

    Private Sub PawnTransAction(sStockID As String, sQuantity As String, sDescription As String, sSubTotal As String, sPrice As String, dtTranDate As Date)
        Dim sLastCmd As String

        Try
            sLastCmd = ""

            '//Process PawnStock Sell//
            '**************************
            'Update to PawnStock
            sSQL = "Update PawnStock set QuantitySold=QuantitySold+" & sQuantity & " where PSID=" & sStockID
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            reader.Close()

            'Update to CashTransactions
            sSQL = "INSERT INTO CashTransactions " & _
                "(CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES " & _
                "('" & p_CashRegister & "','Pawn Stock Sale','" & dtTranDate & "',0," & sStockID & "," & sQuantity & ",0, 0,'" & sDescription & "','" & sSubTotal & "','" & sPrice & "',0,0)"
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            reader2.Close()

            sLastCmd = "cmdInsertCashTransaction" & "," & p_CashRegister & "," & "Pawn Stock Sale" & "," & dtTranDate & "," & 0 & "," & sStockID & "," & sQuantity & "," & 0 & "," & 0 & "," & sDescription & "," & sSubTotal & "," & sPrice & "," & 0 & "," & 0
            'MsgBox(sLastCmd)

            SetCBBIDStatus(sStockID)

            If p_Audit = True Then AuditToFile("Stock Sale", sLastCmd)

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

    Private Sub NormalTransAction(sStockID As String, sQuantity As String, sDescription As String, sSubTotal As String, sPrice As String, dtTranDate As Date, bCellAssesories As Boolean)
        'Update NormalStock set CellphoneAssessories = 2
        Dim sLastCmd As String

        Try
            sLastCmd = ""
            '            cn.BeginTrans()

            '//Process NormalStock Sell//
            '***************************
            'Update to NormalStock
            sSQL = "Update NormalStock set QuantitySold=QuantitySold+" & sQuantity & " where NSID=" & sStockID
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            reader.Close()

            'Update to CashTransactions

            If bCellAssesories Then
                'Standard Stock
                If IsTangoStock(sDescription) = True Then
                    sSQL = "INSERT INTO CashTransactions " & _
                           "(CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES " & _
                           "('" & p_CashRegister & "','Tango Stock Sale','" & dtTranDate & "', 0, 0, " & sQuantity & "," & sStockID & ", 0,'" & sDescription & "','" & sSubTotal & "','" & sPrice & "',0,0)"
                    sLastCmd = "cmdInsertCashTransaction1" & "," & p_CashRegister & "," & "Tango Stock Sale" & "," & dtTranDate & "," & 0 & "," & 0 & "," & sQuantity & "," & sStockID & "," & 0 & "," & sDescription & "," & sSubTotal & "," & sPrice & "," & 0 & "," & 0
                Else
                    sSQL = "INSERT INTO CashTransactions " & _
                           "(CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES " & _
                           "('" & p_CashRegister & "','Standard Stock Sale','" & dtTranDate & "', 0, 0, " & sQuantity & "," & sStockID & ", 0,'" & sDescription & "','" & sSubTotal & "','" & sPrice & "',0,0)"
                    sLastCmd = "cmdInsertCashTransaction2" & "," & p_CashRegister & "," & "Standard Stock Sale" & "," & dtTranDate & "," & 0 & "," & 0 & "," & sQuantity & "," & sStockID & "," & 0 & "," & sDescription & "," & sSubTotal & "," & sPrice & "," & 0 & "," & 0
                End If
            Else
                'Normal Stock
                If IsTangoStock(sDescription) = True Then
                    sSQL = "INSERT INTO CashTransactions " & _
                           "(CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES " & _
                           "('" & p_CashRegister & "','Tango Stock Sale','" & dtTranDate & "', 0, 0, " & sQuantity & "," & sStockID & ", 0,'" & sDescription & "','" & sSubTotal & "','" & sPrice & "',0,0)"
                    sLastCmd = "cmdInsertCashTransaction3" & "," & p_CashRegister & "," & "Tango Stock Sale" & "," & dtTranDate & "," & 0 & "," & 0 & "," & sQuantity & "," & sStockID & "," & 0 & "," & sDescription & "," & sSubTotal & "," & sPrice & "," & 0 & "," & 0
                Else
                    sSQL = "INSERT INTO CashTransactions " & _
                           "(CashRegister, Type, TransactionDate, PNCID, PSID, Quantity, NSID, Category, Description, Amount, UnitPrice, LBID, LBHID) VALUES " & _
                           "('" & p_CashRegister & "','Normal Stock Sale','" & dtTranDate & "', 0, 0, " & sQuantity & "," & sStockID & ", 0,'" & sDescription & "','" & sSubTotal & "','" & sPrice & "',0,0)"
                    sLastCmd = "cmdInsertCashTransaction4" & "," & p_CashRegister & "," & "Normal Stock Sale" & "," & dtTranDate & "," & 0 & "," & 0 & "," & sQuantity & "," & sStockID & "," & 0 & "," & sDescription & "," & sSubTotal & "," & sPrice & "," & 0 & "," & 0
                End If
            End If

            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            reader2.Close()

            If p_Audit = True Then AuditToFile("Stock Sale", sLastCmd)

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

    Private Function IsTangoStock(sStock_Description As String) As Boolean

        If InStr(1, UCase(sStock_Description), "TANGO", vbTextCompare) Then
            IsTangoStock = True
        Else
            IsTangoStock = False
        End If

    End Function

    Private Sub SetCBBIDStatus(sPSID As String)
        Dim sPnCID As String
        Dim arrPnCIDStock(100) As String
        Dim iBuy As Integer
        Dim iSold As Integer
        Dim iPnc As Integer
        Dim iSQL As Integer

        Try

            iPnc = 0
            sPnCID = ""

            sSQL = "Select PnCID from PawnStock where PSID=" & sPSID
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                sPnCID = reader(0).ToString
            End While
            reader.Close()

            sSQL = "Select PSID from PawnStock where PnCID=" & sPnCID
            com.CommandText = sSQL
            Dim reader2 As OleDbDataReader = com.ExecuteReader()
            While reader2.Read()
                arrPnCIDStock(iPnc) = reader2(0).ToString
                iPnc = iPnc + 1
            End While
            reader2.Close()

            For iSQL = 0 To iPnc - 1
                sSQL = "Select Quantity , QuantitySold from PawnStock where PSID=" & arrPnCIDStock(iSQL)
                com.CommandText = sSQL
                Dim reader3 As OleDbDataReader = com.ExecuteReader()
                While reader3.Read()
                    iBuy = CInt(reader3(0).ToString)
                    iSold = CInt(reader3(1).ToString)
                    If iBuy > iSold Then
                        reader3.Close()
                        GoTo lblExit
                    End If
                End While
                reader3.Close()
            Next iSQL

            sSQL = "update PawningTransactions set Status = 2 where PnCID = " & sPnCID
            com.CommandText = sSQL
            Dim reader4 As OleDbDataReader = com.ExecuteReader()
            reader4.Close()

lblExit:

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

End Class
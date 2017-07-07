Imports System.Data.OleDb

Public Class frmRptPawnStock

    Private Sub frmPawnStockReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadList()

    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        With frmReports
            .Label1.Text = "Generating Pawn Report ..."
            .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
            .Height = 67
            .Report = frmReports.ReportType.PawnReport
            .ShowDialog(Me)
        End With
    End Sub

    Private Sub ClearForm()
        Try
            dgPawnStock.ClearSelection()
            If dgPawnStock.Rows.Count <> 0 Then
                Do While dgPawnStock.Rows.Count <> 0
                    dgPawnStock.Rows.Remove(dgPawnStock.Rows(0))
                Loop
            End If
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

    Private Sub LoadList()
        Dim lGroupTotal As Long
        Dim sType As String

        Dim BoldRow As New DataGridViewCellStyle With {.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, FontStyle.Bold)}

        lGroupTotal = 0
        sType = ""

        Try
            dgPawnStock.ClearSelection()

            com.CommandText = "SELECT Status, PSID, Description, SerialNr, Quantity as QuantityTotal, QuantityLayBuy as [Quantity on LayBuy], QuantitySold, (Quantity - QuantityLayBuy - QuantitySold) as [Quantity Available], Price as UnitPrice, ((Quantity - QuantityLayBuy - QuantitySold)*Price) as [Realizable Profit] FROM PawnStock where (((status=1) or (status=3)) and ((Quantity - QuantityLayBuy - QuantitySold)>0)) ORDER BY Status,PSID"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                If sType <> reader(0).ToString() Then
                    If sType <> "" Then 'End of Previous Type...
                        Me.dgPawnStock.Rows.Add("", "", "Status " & sType & " Total:", "", "", _
                                                     "", "", "", "", GetValue(lGroupTotal))
                        Me.dgPawnStock.Rows(dgPawnStock.RowCount - 1).DefaultCellStyle = BoldRow
                        lGroupTotal = 0
                    End If
                    sType = reader(0).ToString()
                    'lGroupTotal = lGroupTotal + reader(9).ToString()

                End If

                Me.dgPawnStock.Rows.Add(reader(0).ToString(), reader(1).ToString(), reader(2).ToString(), reader(3).ToString(), reader(4).ToString(), _
                          reader(5).ToString(), reader(6).ToString(), reader(7).ToString(), GetValue(reader(8).ToString()), GetValue(reader(9).ToString()))
                lGroupTotal = lGroupTotal + reader(9).ToString()

            End While
            reader.Close()

            Me.dgPawnStock.Rows.Add("", "", "Status " & sType & " Total:", "", "", _
                                         "", "", "", "", GetValue(lGroupTotal))


            Me.dgPawnStock.Rows(dgPawnStock.RowCount - 1).DefaultCellStyle = BoldRow

            Me.dgPawnStock.ClearSelection()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        ClearForm()
        Me.Close()

    End Sub

    Private Sub PawnStockBindingSource_CurrentChanged(sender As Object, e As EventArgs) Handles PawnStockBindingSource.CurrentChanged

    End Sub

    Private Sub dgPawnStock_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgPawnStock.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")


        For i = 0 To dgPawnStock.RowCount - 2
            For j = 0 To dgPawnStock.ColumnCount - 1
                For k As Integer = 1 To dgPawnStock.Columns.Count
                    xlWorkSheet.Cells(1, k) = dgPawnStock.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = dgPawnStock(j, i).Value.ToString()
                Next
            Next
        Next

        xlWorkSheet.SaveAs("D:\vbexcel.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file D:\vbexcel.xlsx")
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


End Class
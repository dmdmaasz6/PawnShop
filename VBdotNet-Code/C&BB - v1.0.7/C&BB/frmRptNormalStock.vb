Imports System.Data.OleDb

Public Class frmRptNormalStock

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        With frmReports
            .Label1.Text = "Generating Stock Report ..."
            .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
            .Height = 67
            .Report = frmReports.ReportType.StockReport
            .ShowDialog(Me)
        End With

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        ClearForm()
        Me.Close()

    End Sub


    Private Sub ClearForm()
        Try
            dgNormalStock.ClearSelection()
            If dgNormalStock.Rows.Count <> 0 Then
                Do While dgNormalStock.Rows.Count <> 0
                    dgNormalStock.Rows.Remove(dgNormalStock.Rows(0))
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

    Private Function GetDate(sValIn As String) As String
        If sValIn <> "" Then
            GetDate = FormatDateTime(sValIn, DateFormat.ShortDate)
        Else
            GetDate = "-"
        End If

    End Function

    Private Sub LoadList()
        Dim lGroupTotal As Long
        Dim lBigTotal As Long
        Dim sType As String

        Dim BoldRow As New DataGridViewCellStyle With {.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, FontStyle.Bold)}

        lGroupTotal = 0
        lBigTotal = 0
        sType = ""

        Try
            dgNormalStock.ClearSelection()

            com.CommandText = "SELECT CellphoneAssessories, NSID, DateBought, Description, PurchasePrice, Quantity as QuantityTotal, QuantityLayBuy, QuantitySold, (Quantity - QuantityLayBuy - QuantitySold) as QuantityAvailable, RecommendedSellingPrice, ((Quantity - QuantityLayBuy - QuantitySold)*RecommendedSellingPrice) as RealizableIncome FROM NormalStock where ((Quantity - QuantityLayBuy - QuantitySold)>0) order by CellphoneAssessories,NSID"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(2).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                If sType <> reader(0).ToString() Then
                    If sType <> "" Then 'End of Previous Type...
                        Me.dgNormalStock.Rows.Add("", "", "", "Stock Type " & sType & " Total:", "", "", _
                                                     "", "", "", "", GetValue(lGroupTotal))
                        Me.dgNormalStock.Rows(dgNormalStock.RowCount - 1).DefaultCellStyle = BoldRow
                        lGroupTotal = 0
                    End If
                    sType = reader(0).ToString()
                    'lGroupTotal = lGroupTotal + reader(9).ToString()

                End If

                Me.dgNormalStock.Rows.Add(reader(0).ToString(), reader(1).ToString(), GetDate(reader(2).ToString()), reader(3).ToString(), GetValue(reader(4).ToString()), _
                          reader(5).ToString(), reader(6).ToString(), reader(7).ToString(), reader(8).ToString(), GetValue(reader(9).ToString()), GetValue(reader(10).ToString()))
                lGroupTotal = lGroupTotal + reader(10).ToString()
                lBigTotal = lBigTotal + reader(10).ToString()

            End While
            reader.Close()

            Me.dgNormalStock.Rows.Add("", "", "", "Stock Type " & sType & " Total:", "", "", _
                                         "", "", "", "", GetValue(lGroupTotal))

            Me.dgNormalStock.Rows(dgNormalStock.RowCount - 1).DefaultCellStyle = BoldRow

            Me.dgNormalStock.Rows.Add("", "", "", "Total Value of Normal Stock:", "", "", _
                             "", "", "", "", GetValue(lBigTotal))

            Me.dgNormalStock.Rows(dgNormalStock.RowCount - 1).DefaultCellStyle = BoldRow

            Me.dgNormalStock.ClearSelection()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub frmRptNormalStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadList()

    End Sub
End Class
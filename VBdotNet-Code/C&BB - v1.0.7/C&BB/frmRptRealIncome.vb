Imports System.Data.OleDb

Public Class frmRptRealIncome

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        ClearForm()
        Me.Close()

    End Sub

    Private Sub frmRealIncome_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadList()

    End Sub

    Private Sub ClearForm()
        Try
            dgRealIncome.ClearSelection()
            If dgRealIncome.Rows.Count <> 0 Then
                Do While dgRealIncome.Rows.Count <> 0
                    dgRealIncome.Rows.Remove(dgRealIncome.Rows(0))
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
        Dim BoldRow As New DataGridViewCellStyle With {.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, FontStyle.Bold)}
        Dim sPNCID As String

        lGroupTotal = 0
        sPNCID = ""

        Try
            dgRealIncome.ClearSelection()

            'With Me.dgRealIncome
            '    .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            '    .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            '    .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            '    .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '    .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'End With

            'com.CommandText = "SELECT a.Status, a.PSID, a.Description, a.SerialNr, a.PnCID as PnCID, a.Quantity as QtyTotal, a.QuantitySold, (a.Quantity - a.QuantityLayBuy - a.QuantitySold) as QtyAvail, b.PurchaseAmount as UnitPrice, (QtyAvail * b.BuyBackAmount) as RealProfit FROM PawnStock a RIGHT JOIN PawningTransactions b on b.PnCID = a.PnCID where a.status=1 ORDER BY PSID"
            com.CommandText = "SELECT a.Status, a.PSID, a.Description, a.SerialNr, a.PnCID as PnCID, a.Quantity as QtyTotal, a.QuantitySold, (a.Quantity - a.QuantityLayBuy - a.QuantitySold) as QtyAvail, b.PurchaseAmount as UnitPrice, b.BuyBackAmount as RealProfit FROM PawnStock a RIGHT JOIN PawningTransactions b on b.PnCID = a.PnCID where a.status=1 ORDER BY PSID"
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                If sPNCID = reader(4).ToString() Then
                    Me.dgRealIncome.Rows.Add(reader(1).ToString(), reader(2).ToString(), reader(3).ToString(), reader(4).ToString(), _
                          reader(5).ToString(), reader(6).ToString(), reader(7).ToString(), "-", "-")
                    sPNCID = reader(4).ToString()
                Else
                    Me.dgRealIncome.Rows.Add(reader(1).ToString(), reader(2).ToString(), reader(3).ToString(), reader(4).ToString(), _
                                              reader(5).ToString(), reader(6).ToString(), reader(7).ToString(), GetValue(reader(8).ToString()), GetValue(reader(9).ToString()))
                    sPNCID = reader(4).ToString()
                    lGroupTotal = lGroupTotal + reader(9).ToString()
                End If
            End While
            reader.Close()

            Me.dgRealIncome.Rows.Add("", "Realizable Income Total:", "", "", _
                                          "", "", "", "", GetValue(lGroupTotal))


            Me.dgRealIncome.Rows(dgRealIncome.RowCount - 1).DefaultCellStyle = BoldRow

            Me.dgRealIncome.ClearSelection()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        With frmReports
            .Label1.Text = "Generating Realizable Income Report ..."
            .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
            .Height = 67
            .Report = frmReports.ReportType.RealIncome
            .ShowDialog(Me)
        End With
    End Sub
End Class
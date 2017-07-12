Imports System.Data.OleDb

Public Class frmRptTangoStockCode

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        With frmReports
            .Label1.Text = "Generating Tango Stock Code List ..."
            .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
            .Height = 67
            .Report = frmReports.ReportType.TangoStockList
            .ShowDialog(Me)
        End With

    End Sub

    Private Sub ClearForm()
        Try
            dgTangoStockCode.ClearSelection()
            If dgTangoStockCode.Rows.Count <> 0 Then
                Do While dgTangoStockCode.Rows.Count <> 0
                    dgTangoStockCode.Rows.Remove(dgTangoStockCode.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try


    End Sub

    Private Sub LoadList()
        Try
            dgTangoStockCode.ClearSelection()

            com.CommandText = "SELECT NSID, Description, RecommendedSellingPrice, Quantity FROM NormalStock WHERE (CellphoneAssessories = 2) order by Description"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(2).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                Me.dgTangoStockCode.Rows.Add(reader(0).ToString(), reader(1).ToString())

            End While
            reader.Close()

            Me.dgTangoStockCode.ClearSelection()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        ClearForm()
        Me.Close()

    End Sub

    Private Sub frmRptTangoStockCode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadList()

    End Sub

End Class
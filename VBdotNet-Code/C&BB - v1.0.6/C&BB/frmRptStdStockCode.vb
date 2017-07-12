Imports System.Data.OleDb

Public Class frmRptStdStockCode

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        With frmReports
            .Label1.Text = "Generating Standard Stock Code List ..."
            .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
            .Height = 67
            .Report = frmReports.ReportType.StdStockList
            .ShowDialog(Me)
        End With

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        ClearForm()
        Me.Close()

    End Sub

    Private Sub ClearForm()
        Try
            dgStockCode.ClearSelection()
            If dgStockCode.Rows.Count <> 0 Then
                Do While dgStockCode.Rows.Count <> 0
                    dgStockCode.Rows.Remove(dgStockCode.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try


    End Sub

    Private Sub LoadList()
        Try
            dgStockCode.ClearSelection()

            com.CommandText = "SELECT NSID, Description, RecommendedSellingPrice, Quantity FROM NormalStock WHERE (CellphoneAssessories = 1) order by Description"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Console.WriteLine(reader(0).ToString() + reader(2).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                Me.dgStockCode.Rows.Add(reader(0).ToString(), reader(1).ToString())

            End While
            reader.Close()

            Me.dgStockCode.ClearSelection()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub frmRptStdStockCode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadList()

    End Sub

End Class
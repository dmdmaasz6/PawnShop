Imports System.Data.OleDb

Public Class frmStockSearch

    Private Sub frmStockSearch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ClearList()
        txtSearch.Text = ""
        p_Search_NSID = "'"
        p_Search_PSID = ""

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        cmdSell.Visible = True
        p_Search_NSID = ""
        p_Search_PSID = ""
        Me.Close()

    End Sub

    Private Sub cmdSearcch_Click(sender As Object, e As EventArgs) Handles cmdSearcch.Click
        ClearList()
        LoadList(txtSearch.Text)

    End Sub

    Private Sub ClearList()
        Try
            dgStockSearch.ClearSelection()
            If dgStockSearch.Rows.Count <> 0 Then
                Do While dgStockSearch.Rows.Count <> 0
                    dgStockSearch.Rows.Remove(dgStockSearch.Rows(0))
                Loop
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub LoadList(sDesc As String)
        Dim bValues As Boolean

        Try
            bValues = False

            dgStockSearch.ClearSelection()

            With Me.dgStockSearch
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            com.CommandText = "SELECT Description, SerialNr, (Quantity - QuantityLayBuy - QuantitySold) AS Quantity, Price, ('PSID: ' & PSID) AS ID FROM PawnStock WHERE (((Quantity - QuantityLayBuy - QuantitySold) > 0) AND (status = 3) AND (Description LIKE '%" & sDesc & "%')) UNION SELECT Description, 'No SerialNr', (Quantity - QuantityLayBuy - QuantitySold) AS Quantity, RecommendedSellingPrice AS Price, ('NSID: ' & NSID) AS ID FROM NormalStock WHERE (((Quantity - QuantityLayBuy - QuantitySold) > 0) AND (Description LIKE '%" & sDesc & "%')) ORDER BY Description"

            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                bValues = True
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Me.dgStockSearch.Rows.Add(reader(0).ToString(), reader(1).ToString(), reader(2).ToString(), GetValue(reader(3).ToString()), reader(4).ToString())
            End While
            reader.Close()

            If bValues = False Then
                MsgBox("Sorry, no match was found.", vbInformation, "No match")
            End If

            Me.dgStockSearch.ClearSelection()

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

    Private Sub cmdSell_Click(sender As Object, e As EventArgs) Handles cmdSell.Click
        Dim sSearchId As String

        If Me.dgStockSearch.RowCount = 0 Then
            Exit Sub
        End If

        sSearchId = Me.dgStockSearch.Item(4, dgStockSearch.CurrentRow.Index).Value

        If UCase(Microsoft.VisualBasic.Left(sSearchId, 4)) = "PSID" Then
            p_Search_PSID = Microsoft.VisualBasic.Right(sSearchId, Len(sSearchId) - 5)
            p_Search_NSID = ""
        ElseIf UCase(Microsoft.VisualBasic.Left(sSearchId, 4)) = "NSID" Then
            p_Search_NSID = Microsoft.VisualBasic.Right(sSearchId, Len(sSearchId) - 5)
            p_Search_PSID = ""
        Else
            p_Search_PSID = ""
            p_Search_NSID = ""
        End If

        'MsgBox("Item: " & Me.dgStockSearch.Item(0, dgStockSearch.CurrentRow.Index).Value & vbCrLf & _
        '       "with ID: " & sSearchId & vbCrLf & _
        '       "PS ID: " & p_Search_PSID & vbCrLf & _
        '       "NS ID: " & p_Search_NSID & vbCrLf & _
        '       "has been selected to be sold!", MsgBoxStyle.Information + vbOKOnly)
        Me.Close()

    End Sub
End Class
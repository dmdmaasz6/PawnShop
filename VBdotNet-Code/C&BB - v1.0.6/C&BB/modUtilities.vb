Imports System.Data.OleDb
Imports System.IO

Module modUtilities
    Public con As New Data.OleDb.OleDbConnection()
    Public com As New Data.OleDb.OleDbCommand

    Public arrStockSold(20, 3) As String
    Public iStock As Integer

    Public sTemplatePath As String
    Public sTemplateShop As String


    Public Sub DBConnect(sDBPath As String)

        'Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        '                                                   "Data Source=e:\My Documents\db1.mdb")


        '        Dim con As New Data.OleDb.OleDbConnection()
        con.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPath '"C:\Visual Studio 2010\Projects\Database\dbname.mdb"
        con.Open()
        'Dim com As New Data.OleDb.OleDbCommand
        com.Connection = con
        Exit Sub

    End Sub

    Public Sub DBClose()
        con.Close()

    End Sub

    Public Sub SearchCustomer()

        com.CommandText = "SELECT CustID, IDNr, Name, Address, TelH, TelW FROM Customers WHERE Name LIKE '%BOTHA%'"
        'com.CommandText = "INSERT INTO tableName(fname1,fname2,fname3,fname4) VALUES ('" & txtFname.Text & "','" & txFclass.Text & "','" & txtFgender.Text & "','" & cboGroup.SelectedText & "')"
        'If txtFname.Text = Nothing OrElse txtFclass.Text = Nothing OrElse txtFgender.Text = Nothing Then
        '    MessageBox.Show("Required fields required must be filled!")
        '    Exit Sub
        'Else

        '    MessageBox.Show("Record saved successfully!")
        '    txtFname.Clear()
        '    txtFclass.Clear()
        '    txtGender.Clear()
        '    cboGroup.ResetText()

        'End If

        '       com.ExecuteNonQuery()
        Dim reader As OleDbDataReader = com.ExecuteReader()

        While reader.Read()
            Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
            Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
        End While
        reader.Close()

    End Sub

    Public Sub DailyReport()
        'SELECT CashRegister, Type, TransactionDate, PNCID, PSID, NSID, LBID, LBHID, Description, Category, UnitPrice, Quantity, Amount FROM CashTransactions where TransactionDate = ? and CashRegister = ?

        com.CommandText = "SELECT CashRegister, Type, TransactionDate, PNCID, PSID, NSID, LBID, LBHID, Description, Category, UnitPrice, Quantity, Amount FROM CashTransactions where TransactionDate = #06-06-2013# ORDER by Type"

        Dim reader As OleDbDataReader = com.ExecuteReader()

        While reader.Read()
            Console.WriteLine(reader(0).ToString() + ";" + reader(1).ToString() + ";" + reader(2).ToString() + ";" + reader(3).ToString() + ";" + reader(4).ToString() + ";" + reader(5).ToString() + ";" + reader(6).ToString() + ";" + reader(7).ToString() + ";" + reader(8).ToString() + ";" + reader(9).ToString() + ";" + reader(10).ToString() + ";" + reader(11).ToString() + ";" + reader(12).ToString())
            Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString() + reader(3).ToString() + reader(4).ToString() + reader(5).ToString() + reader(6).ToString() + reader(7).ToString() + reader(8).ToString() + reader(9).ToString() + reader(10).ToString() + reader(11).ToString() + reader(12).ToString())
        End While
        reader.Close()

    End Sub

    Public Sub ErrorToFile(ErrorClass As String, ErrorMethod As String, ErrorMsg As String, SR As String)
        Dim ErrorFile As String
        Dim YY As String
        Dim MM As String
        Dim DD As String

        YY = Year(Now)
        MM = Month(Now)
        If Len(MM) = 1 Then
            MM = "0" & MM
        End If
        DD = Microsoft.VisualBasic.Day(Now)
        If Len(DD) = 1 Then
            DD = "0" & DD
        End If

        ErrorFile = Application.StartupPath & "\Logs\" & Application.ProductName & "_Error_" & YY & MM & DD & ".log"
        'MsgBox(ErrorFile)
        Dim objReader As StreamWriter
        Try
            objReader = File.AppendText(ErrorFile)

            objReader.WriteLine(Now & " : " & ErrorClass & " - " & ErrorMethod)
            objReader.WriteLine("Error Message: " & ErrorMsg)
            objReader.WriteLine(SR)
            objReader.WriteLine("================================================================================================================================================")
            objReader.Close()
            MsgBox("Error message written to file " & vbCrLf & _
                   ErrorFile, vbExclamation + vbOKOnly, "Cash & BuyBack")
        Catch Ex As Exception
            MsgBox("Error wrting to ErrorFile" & vbCrLf & _
                   "Copy error and email to support!" & vbCrLf & vbCrLf & _
                   "Original Error:" & vbCrLf & _
                   ErrorMsg & vbCrLf & _
                   SR & vbCrLf & vbCrLf & _
                   "This Error:" & vbCrLf & _
                   Ex.Message & vbCrLf & _
                   Ex.StackTrace, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Cash & BuyBack Error")
        End Try

    End Sub

    Public Sub AuditToFile(AuditMsg As String, AddAuditMsg As String)
        Dim AuditFile As String
        Dim YY As String
        Dim MM As String
        Dim DD As String

        YY = Year(Now)
        MM = Month(Now)
        If Len(MM) = 1 Then
            MM = "0" & MM
        End If
        DD = Microsoft.VisualBasic.Day(Now)
        If Len(DD) = 1 Then
            DD = "0" & DD
        End If

        AuditFile = Application.StartupPath & "\Logs\" & Application.ProductName & "_Audit_" & YY & MM & DD & ".log"
        'MsgBox(ErrorFile)
        Dim objReader As StreamWriter
        Try
            objReader = File.AppendText(AuditFile)

            objReader.WriteLine(Now & " : " & AuditMsg)
            If AddAuditMsg <> "" Then
                objReader.WriteLine(AddAuditMsg)
            End If
            objReader.Close()
            'MsgBox("Audit message written to file " & vbCrLf & _
            '       AuditFile, vbExclamation + vbOKOnly, "Cash & BuyBack")
        Catch Ex As Exception
            MsgBox("Error wrting to AuditFile" & vbCrLf & _
                   "Copy error and email to support!" & vbCrLf & vbCrLf & _
                   "Original Audit Message:" & vbCrLf & _
                   AuditMsg & vbCrLf & vbCrLf & _
                   "This Error:" & vbCrLf & _
                   Ex.Message & vbCrLf & _
                   Ex.StackTrace, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Cash & BuyBack Error")
        End Try

    End Sub

    Public Sub CheckDate()
        Dim k1 As DateTime = Now

        If k1.CompareTo(p_Date) < 0 Then
            AuditToFile("Date Error: Current System Date is 'older' than last logon date! Current System Date is " & FormatDateTime(Now, DateFormat.LongDate) & ", and last logon date was " & FormatDateTime(p_Date, DateFormat.LongDate) & ".", "")
            MsgBox("Date Error: Current System Date is 'older' than last logon date! " & vbCrLf & _
                    "Current System Date: " & FormatDateTime(Now, DateFormat.LongDate) & vbCrLf & _
                    "Last Logon Date: " & FormatDateTime(p_Date, DateFormat.LongDate), _
                    MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Cash & BuyBack Error")
        End If

    End Sub

    Public Sub ClearStockList()
        Dim i As Integer

        iStock = 0
        For i = 0 To 20
            arrStockSold(i, 0) = ""
            arrStockSold(i, 1) = ""
            arrStockSold(i, 2) = ""
            arrStockSold(i, 3) = ""
        Next i

    End Sub

End Module



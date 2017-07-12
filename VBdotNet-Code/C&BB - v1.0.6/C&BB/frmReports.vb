'Imports Microsoft.Reporting.WinForms
Imports Word = Microsoft.Office.Interop.Word
Imports System.Data.OleDb
Imports System.Threading
Imports Microsoft.Office.Interop.Word

Public Class frmReports
    Public sReportType As String
    Public sDateFrom As String
    Public sDateTo As String
    Private sSQL As String

    Public Enum ReportType
        NoReport
        DailyReport
        CustomReport
        PawnHisotry
        StdStockList
        TangoStockList
        StockReport
        PawnReport
        RealIncome
        StockSale
        ExpiredPawn
    End Enum

    Public Report As ReportType

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        If Report = ReportType.DailyReport Then
            cmdDailyRpt_Click(Nothing, Nothing)
        ElseIf Report = ReportType.CustomReport Then
            cmdCustomRpt_Click(Nothing, Nothing)
        ElseIf Report = ReportType.PawnHisotry Then
            cmdPawnHistRpt_Click(Nothing, Nothing)
        ElseIf Report = ReportType.StdStockList Then
            cmdStdStockList_Click(Nothing, Nothing)
        ElseIf Report = ReportType.TangoStockList Then
            cmdTangoStockLst_Click(Nothing, Nothing)
        ElseIf Report = ReportType.StockReport Then
            cmdStockReport_Click(Nothing, Nothing)
        ElseIf Report = ReportType.PawnReport Then
            cmdPawnReport_Click(Nothing, Nothing)
        ElseIf Report = ReportType.RealIncome Then
            cmdRealIncome_Click(Nothing, Nothing)
        ElseIf Report = ReportType.StockSale Then
            cmdStockSale_Click(Nothing, Nothing)
        ElseIf Report = ReportType.ExpiredPawn Then
            cmdMoveExpiredStock_Click(Nothing, Nothing)
        Else
            MsgBox("Report Type not found!", MsgBoxStyle.Critical + vbOKOnly, "CashBuyBack & CashBuyBack")
            Me.Close()
        End If

        Report = ReportType.NoReport

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

    Private Sub frmDailyReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdDailyRpt_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdDailyRpt.Click

        Dim lTotalAmount As Long
        Dim lTotalQuantity As Long
        Dim lTotalUnitPrice As Long

        Dim lTypeAmount As Long
        Dim lTypeQuantity As Long
        Dim lTypeUnitPrice As Long
        Dim sType As String
        Dim iColumn As Integer
        Dim bVlauesFound As Boolean

        Dim r As Integer, c As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        'Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
        'Dim oRng As Word.Range
        'Dim Pos As Double

        bVlauesFound = False

        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Daily Report of Cash Transactions"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 2 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " " & vbNewLine & _
                                " " & vbNewLine & _
                                "Categorized Expenses: 1=Advertensiekoste; 2=Donasies; 3=Elektrisiteit; 4=Herstelwerk;" & vbNewLine & _
                                "5=Huur; 6=Lone; 7=Petrol en Voertuiguitgawes; 8=Sekuriteitsgelde; 9=Skoonmaakmiddels;" & vbNewLine & _
                                "10=Skryfbehoeftes en Drukwerk; 11=Telefoon; 12=Verversings (Personeel); 13=Allerlei"
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            'Add some text after the table.
            'oTable.Range.InsertParagraphAfter()
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Range.InsertParagraphBefore()
            oPara2.Range.Style = "No Spacing"
            'oPara2.Range.Text = "And here's another table:"

            'Insert x Rows x 10 Coloumns Table for Values
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 10)
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            'oTable.Range.Font.Bold = True
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(1, 1).Range.Text = "CBBID:"
            oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(1.3)
            oTable.Cell(1, 1).Range.Font.Bold = True
            oTable.Cell(1, 2).Range.Text = "PS ID:"
            oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(1.25)
            oTable.Cell(1, 2).Range.Font.Bold = True
            oTable.Cell(1, 3).Range.Text = "NS ID:"
            oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(1.25)
            oTable.Cell(1, 3).Range.Font.Bold = True
            oTable.Cell(1, 4).Range.Text = "LB ID:"
            oTable.Columns.Item(4).Width = oWord.CentimetersToPoints(1.25)
            oTable.Cell(1, 4).Range.Font.Bold = True
            oTable.Cell(1, 5).Range.Text = "LB HID:"
            oTable.Columns.Item(5).Width = oWord.CentimetersToPoints(1.35)
            oTable.Cell(1, 5).Range.Font.Bold = True
            oTable.Cell(1, 6).Range.Text = "Description:"
            oTable.Columns.Item(6).Width = oWord.CentimetersToPoints(5)
            oTable.Cell(1, 6).Range.Font.Bold = True
            oTable.Cell(1, 7).Range.Text = "Category:"
            oTable.Columns.Item(7).Width = oWord.CentimetersToPoints(1.75)
            oTable.Cell(1, 7).Range.Font.Bold = True
            oTable.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 8).Range.Text = "UnitPrice:"
            oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(2)
            oTable.Cell(1, 8).Range.Font.Bold = True
            oTable.Cell(1, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(1, 9).Range.Text = "Quantity:"
            oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(1.75)
            oTable.Cell(1, 9).Range.Font.Bold = True
            oTable.Cell(1, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 10).Range.Text = "Amount:"
            oTable.Columns.Item(10).Width = oWord.CentimetersToPoints(2.15)
            oTable.Cell(1, 10).Range.Font.Bold = True
            oTable.Cell(1, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

            For c = 1 To 10
                oTable.Cell(2, c).Range.Text = ""
            Next

            r = 3
            lTotalAmount = 0
            lTotalQuantity = 0
            lTotalUnitPrice = 0
            lTypeAmount = 0
            lTypeQuantity = 0
            lTypeUnitPrice = 0
            sType = ""

            '        com.CommandText = "SELECT CashRegister, Type, TransactionDate, PNCID, PSID, NSID, LBID, LBHID, Description, Category, UnitPrice, Quantity, Amount FROM CashTransactions where TransactionDate = #" & sDateTo & "# order by Type" 'like '%" & sDateTo & "%' order by Type"
            '        com.CommandText = "SELECT CashRegister, Type, TransactionDate, PNCID, PSID, NSID, LBID, LBHID, Description, Category, UnitPrice, Quantity, Amount FROM CashTransactions where TransactionDate between #" & sDateFrom & "# and #" & sDateTo & "# order by TransactionDate,Type" ' order by Type"
            sSQL = "SELECT CashRegister, Type, TransactionDate, PNCID, PSID, NSID, LBID, LBHID, Description, Category, UnitPrice, Quantity, Amount FROM CashTransactions where TransactionDate = #" & sDateTo & "# order by Type"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                oTable.Rows.Add()
                'MsgBox("CashRegister: " & reader(1).ToString() & vbCrLf & _
                '    "Type: " & reader(1).ToString() & vbCrLf & _
                '    "TransactionDate: " & reader(2).ToString() & vbCrLf & _
                '    "PNCID: " & reader(3).ToString() & vbCrLf & _
                '    "PSID: " & reader(4).ToString() & vbCrLf & _
                '    "NSID: " & reader(5).ToString() & vbCrLf & _
                '    "LBID: " & reader(6).ToString() & vbCrLf & _
                '    "LBHID: " & reader(7).ToString() & vbCrLf & _
                '    "Description: " & reader(8).ToString() & vbCrLf & _
                '    "Category: " & reader(9).ToString() & vbCrLf & _
                '    "UnitPrice: " & reader(10).ToString() & vbCrLf & _
                '    "Quantity: " & reader(11).ToString() & vbCrLf & _
                '    "Amount: " & reader(12).ToString())
                If sType <> reader(1).ToString() Then
                    If sType <> "" Then
                        oTable.Cell(iColumn, 10).Range.Text = GetValue(lTypeAmount)
                        oTable.Cell(iColumn, 10).Range.Font.Bold = True
                        oTable.Cell(iColumn, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                        oTable.Cell(iColumn, 1).Merge(oTable.Cell(iColumn, 5))
                        oTable.Cell(iColumn, 2).Merge(oTable.Cell(iColumn, 3))
                        For c = 1 To 10
                            oTable.Cell(r, c).Range.Text = ""
                        Next
                        r = r + 1
                        oTable.Rows.Add()
                    End If
                    sType = reader(1).ToString()
                    lTypeAmount = 0
                    lTypeQuantity = 0
                    lTypeUnitPrice = 0
                    oTable.Cell(r, 1).Range.Text = "Transaction: " & reader(1).ToString()
                    oTable.Cell(r, 1).Range.Font.Bold = True
                    oTable.Cell(r, 7).Range.Text = "Date: " & FormatDateTime(reader(2).ToString(), DateFormat.LongDate) ', "dd MMMM yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    oTable.Cell(r, 7).Range.Font.Bold = True
                    oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    oTable.Cell(r, 9).Range.Text = "Total:"
                    oTable.Cell(r, 9).Range.Font.Bold = True
                    oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'oTable.Cell(r, 5).Range.Text = r 
                    iColumn = r
                    r = r + 1
                    oTable.Rows.Add()
                End If
                oTable.Cell(r, 1).Range.Text = reader(3).ToString()
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.Text = reader(4).ToString()
                oTable.Cell(r, 3).Range.Text = reader(5).ToString()
                oTable.Cell(r, 4).Range.Text = reader(6).ToString()
                oTable.Cell(r, 5).Range.Text = reader(7).ToString()
                oTable.Cell(r, 6).Range.Text = Replace(Microsoft.VisualBasic.Left(reader(8).ToString(), 25), vbNewLine, " ")
                oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 6).Range.Font.Bold = False
                oTable.Cell(r, 7).Range.Text = reader(9).ToString()
                oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 7).Range.Font.Bold = False
                oTable.Cell(r, 8).Range.Text = GetValue(reader(10).ToString())
                lTypeUnitPrice = lTypeUnitPrice + reader(10).ToString()
                lTotalUnitPrice = lTotalUnitPrice + reader(10).ToString()
                oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 9).Range.Text = reader(11).ToString()
                oTable.Cell(r, 9).Range.Font.Bold = False
                lTypeQuantity = lTypeQuantity + reader(11).ToString()
                lTotalQuantity = lTotalQuantity + reader(11).ToString()
                oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 10).Range.Text = GetValue(reader(12).ToString())
                lTypeAmount = lTypeAmount + reader(12).ToString()
                lTotalAmount = lTotalAmount + reader(12).ToString()
                oTable.Cell(r, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                r = r + 1
            End While
            reader.Close()

            If bVlauesFound = True Then
                oTable.Cell(iColumn, 10).Range.Text = GetValue(lTypeAmount)
                oTable.Cell(iColumn, 10).Range.Font.Bold = True
                oTable.Cell(iColumn, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(iColumn, 1).Merge(oTable.Cell(iColumn, 5))
                oTable.Cell(iColumn, 2).Merge(oTable.Cell(iColumn, 3))

                oTable.Rows.Add()
                For c = 1 To 10
                    oTable.Cell(r, c).Range.Text = ""
                Next
                r = r + 1

                oTable.Rows.Add()
                oTable.Cell(r, 8).Range.Text = "Netto Cash Value for the Day:"
                oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 8).Range.Font.Bold = True
                oTable.Cell(r, 10).Range.Text = GetValue(lTotalAmount)
                oTable.Cell(r, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 10).Range.Font.Bold = True
                oTable.Cell(r, 6).Merge(oTable.Cell(r, 8))
            Else
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 5))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

    Private Sub cmdCustomRpt_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdCustomRpt.Click

        Dim lTotalAmount As Long
        Dim lTotalQuantity As Long
        Dim lTotalUnitPrice As Long

        Dim lTypeAmount As Long
        Dim lTypeQuantity As Long
        Dim lTypeUnitPrice As Long
        Dim sType As String
        Dim iColumn As Integer
        Dim bVlauesFound As Boolean

        Dim r As Integer, c As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        'Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
        'Dim oRng As Word.Range
        'Dim Pos As Double

        bVlauesFound = False

        Try

            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Custom Report of Cash Transactions"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 4 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " " & vbNewLine & _
                                " " & vbNewLine & _
                                "Categorized Expenses: 1=Advertensiekoste; 2=Donasies; 3=Elektrisiteit; 4=Herstelwerk;" & vbNewLine & _
                                "5=Huur; 6=Lone; 7=Petrol en Voertuiguitgawes; 8=Sekuriteitsgelde; 9=Skoonmaakmiddels;" & vbNewLine & _
                                "10=Skryfbehoeftes en Drukwerk; 11=Telefoon; 12=Verversings (Personeel); 13=Allerlei"
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            oTable.Cell(3, 1).Range.Text = ""
            oTable.Cell(4, 1).Merge(oTable.Cell(4, 3))
            oTable.Cell(4, 1).Range.Text = "Custom Report for the period: " & FormatDateTime(sDateFrom, DateFormat.LongDate) & " - " & FormatDateTime(sDateTo, DateFormat.LongDate)
            oTable.Cell(4, 1).Range.Font.Bold = True
            oTable.Cell(4, 1).Range.Font.Color = Word.WdColor.wdColorRed

            'Add some text after the table.
            'oTable.Range.InsertParagraphAfter()
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Range.InsertParagraphBefore()
            oPara2.Range.Style = "No Spacing"
            'oPara2.Range.Text = "And here's another table:"

            'Insert x Rows x 10 Coloumns Table for Values
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 10)
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            'oTable.Range.Font.Bold = True
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(1, 1).Range.Text = "CBBID:"
            oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(1.3)
            oTable.Cell(1, 1).Range.Font.Bold = True
            oTable.Cell(1, 2).Range.Text = "PS ID:"
            oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(1.25)
            oTable.Cell(1, 2).Range.Font.Bold = True
            oTable.Cell(1, 3).Range.Text = "NS ID:"
            oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(1.25)
            oTable.Cell(1, 3).Range.Font.Bold = True
            oTable.Cell(1, 4).Range.Text = "LB ID:"
            oTable.Columns.Item(4).Width = oWord.CentimetersToPoints(1.25)
            oTable.Cell(1, 4).Range.Font.Bold = True
            oTable.Cell(1, 5).Range.Text = "LB HID:"
            oTable.Columns.Item(5).Width = oWord.CentimetersToPoints(1.35)
            oTable.Cell(1, 5).Range.Font.Bold = True
            oTable.Cell(1, 6).Range.Text = "Description:"
            oTable.Columns.Item(6).Width = oWord.CentimetersToPoints(5)
            oTable.Cell(1, 6).Range.Font.Bold = True
            oTable.Cell(1, 7).Range.Text = "Category:"
            oTable.Columns.Item(7).Width = oWord.CentimetersToPoints(1.75)
            oTable.Cell(1, 7).Range.Font.Bold = True
            oTable.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 8).Range.Text = "UnitPrice:"
            oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(2)
            oTable.Cell(1, 8).Range.Font.Bold = True
            oTable.Cell(1, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(1, 9).Range.Text = "Quantity:"
            oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(1.75)
            oTable.Cell(1, 9).Range.Font.Bold = True
            oTable.Cell(1, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 10).Range.Text = "Amount:"
            oTable.Columns.Item(10).Width = oWord.CentimetersToPoints(2.15)
            oTable.Cell(1, 10).Range.Font.Bold = True
            oTable.Cell(1, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

            For c = 1 To 10
                oTable.Cell(2, c).Range.Text = ""
            Next

            r = 3
            lTotalAmount = 0
            lTotalQuantity = 0
            lTotalUnitPrice = 0
            lTypeAmount = 0
            lTypeQuantity = 0
            lTypeUnitPrice = 0
            sType = ""

            'sSQL = "SELECT CashRegister, Type, TransactionDate, PNCID, PSID, NSID, LBID, LBHID, Description, Category, UnitPrice, Quantity, Amount FROM CashTransactions where TransactionDate between #" & sDateFrom & "# and #" & sDateTo & "# order by Type"
            sSQL = "SELECT Type,  Amount FROM CashTransactions where TransactionDate between #" & sDateFrom & "# and #" & sDateTo & "# order by Type"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                oTable.Rows.Add()
                'If sType <> reader(1).ToString() Then
                If sType <> reader(0).ToString() Then
                    If sType <> "" Then
                        oTable.Cell(iColumn, 10).Range.Text = GetValue(lTypeAmount)
                        oTable.Cell(iColumn, 10).Range.Font.Bold = True
                        oTable.Cell(iColumn, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                        oTable.Cell(iColumn, 1).Merge(oTable.Cell(iColumn, 6))
                        For c = 1 To 10
                            oTable.Cell(r, c).Range.Text = ""
                        Next
                        r = r + 1
                        oTable.Rows.Add()
                    End If
                    'sType = reader(1).ToString()
                    sType = reader(0).ToString()
                    lTypeAmount = 0
                    lTypeQuantity = 0
                    lTypeUnitPrice = 0
                    'oTable.Cell(r, 1).Range.Text = "Transaction: " & reader(1).ToString()
                    oTable.Cell(r, 1).Range.Text = "Transaction: " & reader(0).ToString()
                    oTable.Cell(r, 1).Range.Font.Bold = True
                    oTable.Cell(r, 9).Range.Text = "Total:"
                    oTable.Cell(r, 9).Range.Font.Bold = True
                    oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'oTable.Cell(r, 5).Range.Text = r
                    iColumn = r
                    r = r + 1
                End If
                'lTypeUnitPrice = lTypeUnitPrice + reader(10).ToString()
                'lTotalUnitPrice = lTotalUnitPrice + reader(10).ToString()
                'lTypeAmount = lTypeAmount + reader(12).ToString()
                'lTotalAmount = lTotalAmount + reader(12).ToString()
                lTypeAmount = lTypeAmount + reader(1).ToString()
                lTotalAmount = lTotalAmount + reader(1).ToString()

            End While
            reader.Close()

            If bVlauesFound = True Then
                oTable.Cell(iColumn, 10).Range.Text = GetValue(lTypeAmount)
                oTable.Cell(iColumn, 10).Range.Font.Bold = True
                oTable.Cell(iColumn, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(iColumn, 1).Merge(oTable.Cell(iColumn, 5))
                oTable.Cell(iColumn, 2).Merge(oTable.Cell(iColumn, 3))

                oTable.Rows.Add()
                For c = 1 To 10
                    oTable.Cell(r, c).Range.Text = ""
                Next
                r = r + 1

                oTable.Rows.Add()
                oTable.Cell(r, 8).Range.Text = "Netto Cash Value for the Day:"
                oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 8).Range.Font.Bold = True
                oTable.Cell(r, 10).Range.Text = GetValue(lTotalAmount)
                oTable.Cell(r, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 10).Range.Font.Bold = True
                oTable.Cell(r, 6).Merge(oTable.Cell(r, 8))
            Else
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 5))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

    Private Sub cmdPawnHistRpt_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdPawnHistRpt.Click

        Dim bVlauesFound As Boolean

        Dim r As Integer, c As Integer
        Dim arrPawnHist(10000, 9)
        Dim iPawnHist As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oTable2 As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        Dim oPara3 As Word.Paragraph ', oPara4 As Word.Paragraph
        Dim lBuyBackAmount As Long
        Dim lPurchaseAmount As Long

        Try

            lBuyBackAmount = 0
            lPurchaseAmount = 0

            For r = 0 To 10000
                For c = 0 To 9
                    arrPawnHist(r, c) = ""
                Next
            Next

            iPawnHist = 0

            bVlauesFound = False

            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            '        oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Pawning History"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 7 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " " & vbNewLine & _
                                " " & vbNewLine & _
                                "Status:"
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(2, 1).Range.Font.Bold = True

            oTable.Cell(3, 1).Merge(oTable.Cell(3, 3))
            oTable.Cell(3, 1).Range.Text = "     1 = Transaction is still waiting to be brought back"
            oTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            oTable.Cell(4, 1).Merge(oTable.Cell(4, 3))
            oTable.Cell(4, 1).Range.Text = "     2 = Transaction has already been brought back"
            oTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            oTable.Cell(5, 1).Merge(oTable.Cell(5, 3))
            oTable.Cell(5, 1).Range.Text = "     3 = Transaction has expired and pawned stock has been moved to the floor to be sold."
            oTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            oTable.Cell(6, 1).Range.Text = ""
            oTable.Cell(7, 1).Merge(oTable.Cell(7, 3))
            oTable.Cell(7, 1).Range.Text = "Pawning History for the period: " & FormatDateTime(sDateFrom, DateFormat.LongDate) & " - " & FormatDateTime(sDateTo, DateFormat.LongDate)
            oTable.Cell(7, 1).Range.Font.Bold = True
            oTable.Cell(7, 1).Range.Font.Color = Word.WdColor.wdColorRed

            'Add some text after the table.
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Range.Style = "No Spacing"

            sSQL = "SELECT PawningTransactions.DateHandedIn, PawningTransactions.PnCID, Customers.Name, Customers.IDNr, PawningTransactions.PurchaseAmount, PawningTransactions.BuyBackAmount, PawningTransactions.ExpiryDate, PawningTransactions.BuyBackDate, PawningTransactions.Status FROM PawningTransactions INNER JOIN Customers ON (PawningTransactions.CustID = Customers.CustID) where (PawningTransactions.DateHandedIn between #" & sDateFrom & "# and #" & sDateTo & "#) order by PawningTransactions.DateHandedIn"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                iPawnHist = iPawnHist + 1
                'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'DateHandedIn
                If reader(0).ToString() <> "" Then
                    arrPawnHist(iPawnHist, 1) = FormatDateTime(reader(0).ToString(), DateFormat.LongDate)
                Else
                    arrPawnHist(iPawnHist, 1) = "-"
                End If
                'BuyBackAmount
                arrPawnHist(iPawnHist, 2) = GetValue(reader(5).ToString())
                lBuyBackAmount = lBuyBackAmount + reader(5).ToString()
                'CBB ID
                arrPawnHist(iPawnHist, 3) = reader(1).ToString()
                'ExpiryDate
                If reader(6).ToString() <> "" Then
                    arrPawnHist(iPawnHist, 4) = FormatDateTime(reader(6).ToString(), DateFormat.LongDate)
                Else
                    arrPawnHist(iPawnHist, 4) = "-"
                End If
                'PurchaseAmount
                arrPawnHist(iPawnHist, 5) = GetValue(reader(4).ToString())
                lPurchaseAmount = lPurchaseAmount + reader(4).ToString()
                'BuyBackDate
                If reader(7).ToString() <> "" Then
                    arrPawnHist(iPawnHist, 6) = FormatDateTime(reader(7).ToString(), DateFormat.LongDate)
                Else
                    arrPawnHist(iPawnHist, 6) = "-"
                End If

                'Name
                arrPawnHist(iPawnHist, 7) = reader(2).ToString()
                'ID Nr
                arrPawnHist(iPawnHist, 8) = reader(3).ToString()
                'Status
                arrPawnHist(iPawnHist, 9) = reader(8).ToString()

            End While
            reader.Close()

            For i = 1 To iPawnHist
                'Insert 6 Rows x 6 Coloumns Table for Values
                oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 6)
                oTable.Range.Style = "No Spacing"
                oTable.Range.Font.Size = 9
                'oTable.Range.Font.Bold = True
                'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
                'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
                'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                oTable.Cell(1, 1).Range.Text = "DateHandedIn:"
                oTable.Cell(1, 1).Range.Font.Bold = True
                oTable.Cell(1, 2).Merge(oTable.Cell(1, 3))
                oTable.Cell(1, 2).Range.Text = arrPawnHist(i, 1)
                oTable.Cell(1, 3).Range.Text = "BuyBackAmount:"
                oTable.Cell(1, 3).Range.Font.Bold = True
                oTable.Cell(1, 4).Range.Text = arrPawnHist(i, 2)
                oTable.Cell(1, 5).Range.Text = " "

                oTable.Cell(2, 1).Range.Text = "CBB ID:"
                oTable.Cell(2, 1).Range.Font.Bold = True
                oTable.Cell(2, 2).Merge(oTable.Cell(2, 3))
                oTable.Cell(2, 2).Range.Text = arrPawnHist(i, 3)
                oTable.Cell(2, 3).Range.Text = "ExpiryDate:"
                oTable.Cell(2, 3).Range.Font.Bold = True
                oTable.Cell(2, 4).Range.Text = arrPawnHist(i, 4)
                oTable.Cell(2, 5).Range.Text = " "

                oTable.Cell(3, 1).Range.Text = "PurchaseAmount:"
                oTable.Cell(3, 1).Range.Font.Bold = True
                oTable.Cell(3, 2).Merge(oTable.Cell(3, 3))
                oTable.Cell(3, 2).Range.Text = arrPawnHist(i, 5)
                oTable.Cell(3, 3).Range.Text = "BuyBackDate:"
                oTable.Cell(3, 3).Range.Font.Bold = True
                oTable.Cell(3, 4).Range.Text = arrPawnHist(i, 6)
                oTable.Cell(3, 5).Range.Text = " "

                oTable.Cell(4, 1).Range.Text = " "
                oTable.Cell(4, 2).Merge(oTable.Cell(4, 3))
                oTable.Cell(4, 2).Range.Text = " "
                oTable.Cell(4, 3).Range.Text = " "
                oTable.Cell(4, 4).Range.Text = " "
                oTable.Cell(4, 5).Range.Text = " "

                oTable.Cell(5, 1).Range.Text = "Name:"
                oTable.Cell(5, 1).Range.Font.Bold = True
                oTable.Cell(5, 2).Merge(oTable.Cell(5, 3))
                oTable.Cell(5, 2).Range.Text = arrPawnHist(i, 7)
                oTable.Cell(5, 3).Range.Text = " "
                oTable.Cell(5, 4).Range.Text = " "
                oTable.Cell(5, 5).Range.Text = " "

                oTable.Cell(6, 1).Range.Text = "ID Nr:"
                oTable.Cell(6, 1).Range.Font.Bold = True
                oTable.Cell(6, 2).Merge(oTable.Cell(6, 3))
                oTable.Cell(6, 2).Range.Text = arrPawnHist(i, 8)
                oTable.Cell(6, 3).Range.Text = "Status:"
                oTable.Cell(6, 3).Range.Font.Bold = True
                oTable.Cell(6, 4).Range.Text = arrPawnHist(i, 9)
                oTable.Cell(6, 5).Range.Text = " "

                oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                'oPara2.Range.InsertParagraphBefore()
                oPara2.Range.Style = "No Spacing"
                oTable2 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 6)
                oTable2.Range.Style = "No Spacing"
                oTable2.Range.Font.Size = 9
                'oTable2.Range.Font.Bold = True
                'oTable2.Borders.OutsideColor = Word.WdColor.wdColorBlack
                'oTable2.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                'oTable2.Borders.InsideColor = Word.WdColor.wdColorBlack
                'oTable2.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                oTable2.Cell(1, 1).Range.Text = "PS ID:"
                oTable2.Cell(1, 1).Range.Font.Bold = True
                oTable2.Columns.Item(1).Width = oWord.CentimetersToPoints(1.5)
                oTable2.Cell(1, 2).Range.Text = "Description:"
                oTable2.Cell(1, 2).Range.Font.Bold = True
                oTable2.Columns.Item(2).Width = oWord.CentimetersToPoints(7.9)
                oTable2.Cell(1, 3).Range.Text = "SerialNr:"
                oTable2.Cell(1, 3).Range.Font.Bold = True
                oTable2.Columns.Item(3).Width = oWord.CentimetersToPoints(3.5)
                oTable2.Cell(1, 4).Range.Text = "Qty:"
                oTable2.Cell(1, 4).Range.Font.Bold = True
                oTable2.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable2.Columns.Item(4).Width = oWord.CentimetersToPoints(2)
                oTable2.Cell(1, 5).Range.Text = "Qty-LayBuy:"
                oTable2.Cell(1, 5).Range.Font.Bold = True
                oTable2.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable2.Columns.Item(5).Width = oWord.CentimetersToPoints(2)
                oTable2.Cell(1, 6).Range.Text = "Qty-Sold:"
                oTable2.Cell(1, 6).Range.Font.Bold = True
                oTable2.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable2.Columns.Item(6).Width = oWord.CentimetersToPoints(2)

                sSQL = "SELECT PnCID, PSID, Description, SerialNr, Quantity, QuantityLayBuy, QuantitySold FROM PawnStock where PnCID = " & arrPawnHist(i, 3)
                com.CommandText = sSQL
                reader = com.ExecuteReader()
                r = 1
                While reader.Read()
                    r = r + 1
                    oTable2.Rows.Add()
                    oTable2.Cell(r, 1).Range.Text = reader(1).ToString
                    oTable2.Cell(r, 1).Range.Font.Bold = False
                    oTable2.Cell(r, 2).Range.Text = Microsoft.VisualBasic.Left(reader(2).ToString, 50)
                    oTable2.Cell(r, 2).Range.Font.Bold = False
                    oTable2.Cell(r, 3).Range.Text = Microsoft.VisualBasic.Left(reader(3).ToString, 19)
                    oTable2.Cell(r, 3).Range.Font.Bold = False
                    oTable2.Cell(r, 4).Range.Text = reader(4).ToString
                    oTable2.Cell(r, 4).Range.Font.Bold = False
                    oTable2.Cell(r, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    oTable2.Cell(r, 5).Range.Text = reader(5).ToString
                    oTable2.Cell(r, 5).Range.Font.Bold = False
                    oTable2.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    oTable2.Cell(r, 6).Range.Text = reader(6).ToString
                    oTable2.Cell(r, 6).Range.Font.Bold = False
                    oTable2.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                    'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                End While
                reader.Close()
                r = r + 1
                oTable2.Rows.Add()
                oTable2.Cell(r, 1).Range.Text = " "
                oTable2.Cell(r, 1).Range.Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                oTable2.Cell(r, 2).Range.Text = " "
                oTable2.Cell(r, 2).Range.Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                oTable2.Cell(r, 3).Range.Text = " "
                oTable2.Cell(r, 3).Range.Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                oTable2.Cell(r, 4).Range.Text = " "
                oTable2.Cell(r, 4).Range.Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                oTable2.Cell(r, 5).Range.Text = " "
                oTable2.Cell(r, 5).Range.Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                oTable2.Cell(r, 6).Range.Text = " "
                oTable2.Cell(r, 6).Range.Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                'Add some text after the table.
                oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                oPara2.Range.Style = "No Spacing"
            Next

            If bVlauesFound = True Then
                oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                oPara2.Range.Style = "No Spacing"
                'Insert 2 Rows x 4 Coloumns Table for Values
                oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 2)
                oTable.Range.Style = "No Spacing"
                oTable.Range.Font.Size = 9
                oTable.Cell(1, 1).Merge(oTable.Cell(1, 2))
                oTable.Cell(1, 1).Range.Text = "Totals as on Report Date: " & Format(Now, "dd MMMM yyyy")
                oTable.Cell(1, 1).Range.Font.Bold = False

                oTable.Cell(2, 1).Range.Text = "PurchaseAmount: " & GetValue(lPurchaseAmount)
                oTable.Cell(2, 1).Select()
                oWord.Selection.Font.Bold = 0
                oWord.Selection.TypeText("PurchaseAmount: ")
                oWord.Selection.Font.Bold = 1
                oWord.Selection.TypeText(GetValue(lPurchaseAmount))
                oTable.Cell(2, 2).Range.Text = "BuyBackAmount: " & GetValue(lBuyBackAmount)
                oTable.Cell(2, 2).Select()
                oWord.Selection.Font.Bold = 0
                oWord.Selection.TypeText("BuyBackAmount: ")
                oWord.Selection.Font.Bold = 1
                oWord.Selection.TypeText(GetValue(lBuyBackAmount))

            Else
                oPara3 = oDoc.Content.Paragraphs.Add
                oPara3.Range.Text = "No Values Found."
                oPara3.Range.Font.Bold = True
                oPara3.Range.Font.Size = 9
                oPara3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

    Private Sub cmdStdStockList_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdStdStockList.Click


        Dim bVlauesFound As Boolean

        Dim r As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph

        bVlauesFound = False

        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Standard Stock Code List"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 1 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

            'Add some text after the table.
            'oTable.Range.InsertParagraphAfter()
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Range.InsertParagraphBefore()
            oPara2.Range.Style = "No Spacing"
            'oPara2.Range.Text = "And here's another table:"

            'Insert x Rows x 2 Coloumns Table for Values
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 2)
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            'oTable.Range.Font.Bold = True
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(1, 1).Range.Text = "NS ID:"
            oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(2)
            oTable.Cell(1, 1).Range.Font.Bold = True
            oTable.Cell(1, 2).Range.Text = "Description:"
            oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(15)
            oTable.Cell(1, 2).Range.Font.Bold = True


            r = 2

            sSQL = "SELECT NSID, Description, RecommendedSellingPrice, Quantity FROM NormalStock WHERE (CellphoneAssessories = 1) order by Description"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = reader(0).ToString()
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.Text = reader(1).ToString()
                oTable.Cell(r, 2).Range.Font.Bold = False
                r = r + 1
            End While
            reader.Close()

            If bVlauesFound <> True Then
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 2))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

    Private Sub cmdTangoStockLst_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdTangoStockLst.Click


        Dim bVlauesFound As Boolean

        Dim r As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph

        bVlauesFound = False

        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Tango Stock Code List"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 1 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

            'Add some text after the table.
            'oTable.Range.InsertParagraphAfter()
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Range.InsertParagraphBefore()
            oPara2.Range.Style = "No Spacing"
            'oPara2.Range.Text = "And here's another table:"

            'Insert x Rows x 2 Coloumns Table for Values
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 2)
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            'oTable.Range.Font.Bold = True
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(1, 1).Range.Text = "NS ID:"
            oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(2)
            oTable.Cell(1, 1).Range.Font.Bold = True
            oTable.Cell(1, 2).Range.Text = "Description:"
            oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(15)
            oTable.Cell(1, 2).Range.Font.Bold = True


            r = 2

            sSQL = "SELECT NSID, Description, RecommendedSellingPrice, Quantity FROM NormalStock WHERE (CellphoneAssessories = 2) order by Description"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = reader(0).ToString()
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.Text = reader(1).ToString()
                oTable.Cell(r, 2).Range.Font.Bold = False
                r = r + 1
            End While
            reader.Close()

            If bVlauesFound <> True Then
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 2))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

    Private Sub cmdStockReport_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdStockReport.Click


        Dim bVlauesFound As Boolean

        Dim r As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph

        Dim lGroupTotal As Long
        Dim lBigTotal As Long
        Dim sType As String

        sType = ""
        lGroupTotal = 0
        lBigTotal = 0

        bVlauesFound = False

        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Report on Stock"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 4 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9

            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " "
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(3, 1).Merge(oTable.Cell(3, 3))
            oTable.Cell(3, 1).Range.Text = "Stock Types:"
            oTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(3, 1).Range.Bold = True
            oTable.Cell(4, 1).Merge(oTable.Cell(4, 3))
            oTable.Cell(4, 1).Range.Text = vbTab & "0 = Normal Stock" & vbNewLine & _
                                vbTab & "1 = Standard Stock" & vbNewLine & _
                                vbTab & "2 = Tango Stock"
            oTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            sSQL = "SELECT CellphoneAssessories, NSID, DateBought, Description, PurchasePrice, Quantity as QuantityTotal, QuantityLayBuy, QuantitySold, (Quantity - QuantityLayBuy - QuantitySold) as QuantityAvailable, RecommendedSellingPrice, ((Quantity - QuantityLayBuy - QuantitySold)*RecommendedSellingPrice) as RealizableIncome FROM NormalStock where ((Quantity - QuantityLayBuy - QuantitySold)>0) order by CellphoneAssessories,NSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()

                bVlauesFound = True
                'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                If sType <> reader(0).ToString() Then
                    If sType <> "" Then 'End of Previous Type...
                        oTable.Rows.Add()
                        oTable.Cell(r, 1).Range.Text = ""
                        oTable.Cell(r, 2).Range.Text = ""
                        oTable.Cell(r, 3).Range.Text = ""
                        oTable.Cell(r, 4).Range.Text = ""
                        oTable.Cell(r, 5).Range.Text = ""
                        oTable.Cell(r, 6).Range.Text = ""
                        oTable.Cell(r, 7).Range.Text = ""
                        oTable.Cell(r, 8).Range.Text = ""
                        oTable.Cell(r, 9).Range.Text = ""
                        oTable.Cell(r, 10).Range.Text = ""
                        oTable.Cell(r, 8).Merge(oTable.Cell(r, 9))

                        oTable.Cell(r, 8).Range.Text = "Group Total:"
                        oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        oTable.Cell(r, 8).Range.Font.Bold = True
                        oTable.Cell(r, 9).Range.Text = GetValue(lGroupTotal)
                        oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                        oTable.Cell(r, 9).Range.Font.Bold = True

                        lBigTotal = lBigTotal + lGroupTotal

                        oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Cell(3, 8).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 8).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                        'oTable.Rows(r).Range.InsertBreak(Word.WdBreakType.wdPageBreak)
                        oDoc.Bookmarks.Item("\endofdoc").Range.InsertBreak(Word.WdBreakType.wdPageBreak)

                        r = r + 1

                    End If
                    sType = reader(0).ToString()
                    lGroupTotal = 0

                    'Add some text after the table.
                    'oTable.Range.InsertParagraphAfter()
                    oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    'oPara2.Range.InsertParagraphBefore()
                    oPara2.Range.Style = "No Spacing"
                    oPara2.Range.Font.Size = 9
                    oPara2.Range.Font.Bold = True

                    'Insert x Rows x 10 Coloumns Table for Values
                    oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 10)
                    oTable.Range.Style = "No Spacing"
                    oTable.Range.Font.Size = 9
                    'oTable.Range.Font.Bold = True
                    'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
                    'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
                    'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle

                    STCKRPT_Headings(oTable, oWord, "Stock Type: " & sType)

                    r = 4

                End If

                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = reader(1).ToString()
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                If reader(2).ToString() <> "" Then
                    oTable.Cell(r, 2).Range.Text = FormatDateTime(reader(2).ToString(), DateFormat.ShortDate)
                Else
                    oTable.Cell(r, 2).Range.Text = "-"
                End If
                oTable.Cell(r, 2).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 3).Range.Text = reader(3).ToString()
                oTable.Cell(r, 3).Range.Font.Bold = False
                oTable.Cell(r, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 4).Range.Text = GetValue(reader(4).ToString())
                oTable.Cell(r, 4).Range.Font.Bold = False
                oTable.Cell(r, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 5).Range.Text = reader(5).ToString()
                oTable.Cell(r, 5).Range.Font.Bold = False
                oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 6).Range.Text = reader(6).ToString()
                oTable.Cell(r, 6).Range.Font.Bold = False
                oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 7).Range.Text = reader(7).ToString()
                oTable.Cell(r, 7).Range.Font.Bold = False
                oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 8).Range.Text = reader(8).ToString()
                oTable.Cell(r, 8).Range.Font.Bold = False
                oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 9).Range.Text = GetValue(reader(9).ToString())
                oTable.Cell(r, 9).Range.Font.Bold = False
                oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 10).Range.Text = GetValue(reader(10).ToString())
                lGroupTotal = lGroupTotal + reader(10).ToString()
                oTable.Cell(r, 10).Range.Font.Bold = False
                oTable.Cell(r, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                r = r + 1
            End While
            reader.Close()

            oTable.Rows.Add()
            oTable.Cell(r, 1).Range.Text = ""
            oTable.Cell(r, 2).Range.Text = ""
            oTable.Cell(r, 3).Range.Text = ""
            oTable.Cell(r, 4).Range.Text = ""
            oTable.Cell(r, 5).Range.Text = ""
            oTable.Cell(r, 6).Range.Text = ""
            oTable.Cell(r, 7).Range.Text = ""
            oTable.Cell(r, 8).Range.Text = ""
            oTable.Cell(r, 9).Range.Text = ""
            oTable.Cell(r, 10).Range.Text = ""
            oTable.Cell(r, 8).Merge(oTable.Cell(r, 9))

            oTable.Cell(r, 8).Range.Text = "Group Total:"
            oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(r, 8).Range.Font.Bold = True
            oTable.Cell(r, 9).Range.Text = GetValue(lGroupTotal)
            oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(r, 9).Range.Font.Bold = True

            lBigTotal = lBigTotal + lGroupTotal
            r = r + 1
            oTable.Rows.Add()
            oTable.Cell(r, 1).Range.Text = ""
            oTable.Cell(r, 2).Range.Text = ""
            oTable.Cell(r, 3).Range.Text = ""
            oTable.Cell(r, 4).Range.Text = ""
            oTable.Cell(r, 5).Range.Text = ""
            oTable.Cell(r, 6).Range.Text = ""
            oTable.Cell(r, 7).Range.Text = ""
            oTable.Cell(r, 8).Range.Text = ""
            oTable.Cell(r, 9).Range.Text = ""
            r = r + 1
            oTable.Rows.Add()
            oTable.Cell(r, 1).Range.Text = ""
            oTable.Cell(r, 2).Range.Text = ""
            oTable.Cell(r, 3).Range.Text = ""
            oTable.Cell(r, 4).Range.Text = ""
            oTable.Cell(r, 5).Range.Text = ""
            oTable.Cell(r, 6).Range.Text = ""
            oTable.Cell(r, 7).Range.Text = ""
            oTable.Cell(r, 8).Range.Text = ""
            oTable.Cell(r, 9).Range.Text = ""
            oTable.Cell(r, 5).Merge(oTable.Cell(r, 8))

            oTable.Cell(r, 5).Range.Text = "Total Value of Normal Stock:"
            oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(r, 5).Range.Font.Bold = True
            oTable.Cell(r, 6).Range.Text = GetValue(lBigTotal)
            oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(r, 6).Range.Font.Bold = True

            oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(3, 8).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 8).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle

            If bVlauesFound <> True Then
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 2))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try
    End Sub

    Public Sub STCKRPT_Headings(oTable As Object, oWord As Object, Heading As String)

        oTable.Cell(1, 1).Range.Text = Heading
        oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(1.25)
        oTable.Cell(1, 1).Range.Font.Bold = True
        oTable.Cell(1, 2).Range.Text = ""
        oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(2.25)
        oTable.Cell(1, 2).Range.Font.Bold = True
        oTable.Cell(1, 3).Range.Text = ""
        oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(4.5)
        oTable.Cell(1, 3).Range.Font.Bold = True
        oTable.Cell(1, 4).Range.Text = ""
        'oTable.Columns.Item(4).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(1, 4).Range.Font.Bold = True
        oTable.Cell(1, 5).Range.Text = ""
        oTable.Columns.Item(5).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(1, 5).Range.Font.Bold = True
        oTable.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        oTable.Cell(1, 6).Range.Text = ""
        oTable.Columns.Item(6).Width = oWord.CentimetersToPoints(1.4)
        oTable.Cell(1, 6).Range.Font.Bold = True
        oTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(1, 7).Range.Text = ""
        oTable.Columns.Item(7).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(1, 7).Range.Font.Bold = True
        oTable.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(1, 8).Range.Text = ""
        oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(1, 8).Range.Font.Bold = True
        oTable.Cell(1, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(1, 9).Range.Text = ""
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(1, 9).Range.Font.Bold = True
        oTable.Cell(1, 10).Range.Text = ""
        'oTable.Columns.Item(10).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(1, 10).Range.Font.Bold = True

        oTable.Rows.Add()

        oTable.Cell(2, 1).Range.Text = ""
        oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(1.25)
        oTable.Cell(2, 1).Range.Font.Bold = True
        oTable.Cell(2, 2).Range.Text = ""
        oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(2.25)
        oTable.Cell(2, 2).Range.Font.Bold = True
        oTable.Cell(2, 3).Range.Text = ""
        oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(4.5)
        oTable.Cell(2, 3).Range.Font.Bold = True
        oTable.Cell(2, 4).Range.Text = ""
        'oTable.Columns.Item(4).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(2, 4).Range.Font.Bold = True
        oTable.Cell(2, 5).Range.Text = "Quantities:"
        oTable.Columns.Item(5).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(2, 5).Range.Font.Bold = True
        oTable.Cell(2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(2, 6).Range.Text = ""
        oTable.Columns.Item(6).Width = oWord.CentimetersToPoints(1.4)
        oTable.Cell(2, 6).Range.Font.Bold = True
        oTable.Cell(2, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(2, 7).Range.Text = ""
        oTable.Columns.Item(7).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(2, 7).Range.Font.Bold = True
        oTable.Cell(2, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(2, 8).Range.Text = ""
        oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(2, 8).Range.Font.Bold = True
        oTable.Cell(2, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(2, 9).Range.Text = ""
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(2, 9).Range.Font.Bold = True
        oTable.Cell(2, 10).Range.Text = ""
        'oTable.Columns.Item(10).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(2, 10).Range.Font.Bold = True

        oTable.Rows.Add()

        oTable.Cell(3, 1).Range.Text = "NS ID:"
        oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(1.25)
        oTable.Cell(3, 1).Range.Font.Bold = True
        oTable.Cell(3, 2).Range.Text = "Date Bought:"
        oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(2.25)
        oTable.Cell(3, 2).Range.Font.Bold = True
        oTable.Cell(3, 3).Range.Text = "Description:"
        oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(4.5)
        oTable.Cell(3, 3).Range.Font.Bold = True
        oTable.Cell(3, 4).Range.Text = "Purchase Price:"
        oTable.Columns.Item(4).Width = oWord.CentimetersToPoints(1.75)
        oTable.Cell(3, 4).Range.Font.Bold = True
        oTable.Cell(3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(3, 5).Range.Text = "Total:"
        oTable.Columns.Item(5).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(3, 5).Range.Font.Bold = True
        oTable.Cell(3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(3, 6).Range.Text = "LayBy:"
        oTable.Columns.Item(6).Width = oWord.CentimetersToPoints(1.4)
        oTable.Cell(3, 6).Range.Font.Bold = True
        oTable.Cell(3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(3, 7).Range.Text = "Sold:"
        oTable.Columns.Item(7).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(3, 7).Range.Font.Bold = True
        oTable.Cell(3, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(3, 8).Range.Text = "Avail.:"
        oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(1.3)
        oTable.Cell(3, 8).Range.Font.Bold = True
        oTable.Cell(3, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(3, 9).Range.Text = "Recomm. Selling Price:"
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(3, 9).Range.Font.Bold = True
        oTable.Cell(3, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Cell(3, 10).Range.Text = "Realizable Income:"
        'oTable.Columns.Item(10).Width = oWord.CentimetersToPoints(2)
        oTable.Cell(3, 10).Range.Font.Bold = True
        oTable.Cell(3, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

        oTable.Cell(1, 1).Merge(oTable.Cell(1, 3))
        oTable.Cell(1, 6).Merge(oTable.Cell(1, 7))
        oTable.Cell(2, 5).Merge(oTable.Cell(2, 8))

    End Sub

    Private Sub cmdPawnReport_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdPawnReport.Click


        Dim bVlauesFound As Boolean

        Dim r As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph

        Dim lGroupTotal As Long
        Dim lBigTotal As Long
        Dim sType As String

        sType = ""
        lGroupTotal = 0
        lBigTotal = 0

        bVlauesFound = False

        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Report on Pawn Stock"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 4 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9

            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " "
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(3, 1).Merge(oTable.Cell(3, 3))
            oTable.Cell(3, 1).Range.Text = "Status:"
            oTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(3, 1).Range.Bold = True
            oTable.Cell(4, 1).Merge(oTable.Cell(4, 3))
            oTable.Cell(4, 1).Range.Text = vbTab & "1 = Pawned Stock waiting to be bought back" & vbNewLine & _
                                vbTab & "3 = Pawned Stock expired and moved to floor to be sold."
            oTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            'sSQL = "SELECT CellphoneAssessories, NSID, DateBought, Description, PurchasePrice, Quantity as QuantityTotal, QuantityLayBuy, QuantitySold, (Quantity - QuantityLayBuy - QuantitySold) as QuantityAvailable, RecommendedSellingPrice, ((Quantity - QuantityLayBuy - QuantitySold)*RecommendedSellingPrice) as RealizableIncome FROM NormalStock where ((Quantity - QuantityLayBuy - QuantitySold)>0) order by CellphoneAssessories,NSID"
            sSQL = "SELECT Status, PSID, Description, SerialNr, Quantity as QuantityTotal, QuantityLayBuy as [Quantity on LayBuy], QuantitySold, (Quantity - QuantityLayBuy - QuantitySold) as [Quantity Available], Price as UnitPrice, ((Quantity - QuantityLayBuy - QuantitySold)*Price) as [Realizable Profit] FROM PawnStock where (((status=1) or (status=3)) and ((Quantity - QuantityLayBuy - QuantitySold)>0)) ORDER BY Status,PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                If sType <> reader(0).ToString() Then
                    If sType <> "" Then 'End of Previous Type...
                        oTable.Rows.Add()
                        oTable.Cell(r, 1).Range.Text = ""
                        oTable.Cell(r, 2).Range.Text = ""
                        oTable.Cell(r, 3).Range.Text = ""
                        oTable.Cell(r, 4).Range.Text = ""
                        oTable.Cell(r, 5).Range.Text = ""
                        oTable.Cell(r, 6).Range.Text = ""
                        oTable.Cell(r, 7).Range.Text = ""
                        oTable.Cell(r, 8).Range.Text = ""
                        oTable.Cell(r, 9).Range.Text = ""
                        oTable.Cell(r, 7).Merge(oTable.Cell(r, 8))

                        oTable.Cell(r, 7).Range.Text = "Status Total:"
                        oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        oTable.Cell(r, 7).Range.Font.Bold = True
                        oTable.Cell(r, 8).Range.Text = GetValue(lGroupTotal)
                        oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                        oTable.Cell(r, 8).Range.Font.Bold = True

                        lBigTotal = lBigTotal + lGroupTotal

                        oTable.Cell(3, 4).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 4).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
                        oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                        'oTable.Rows(r).Range.InsertBreak(Word.WdBreakType.wdPageBreak)
                        oDoc.Bookmarks.Item("\endofdoc").Range.InsertBreak(Word.WdBreakType.wdPageBreak)

                        r = r + 1

                    End If
                    sType = reader(0).ToString()
                    lGroupTotal = 0

                    'Add some text after the table.
                    'oTable.Range.InsertParagraphAfter()
                    oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
                    'oPara2.Range.InsertParagraphBefore()
                    oPara2.Range.Style = "No Spacing"
                    oPara2.Range.Font.Size = 9
                    oPara2.Range.Font.Bold = True

                    'Insert x Rows x 10 Coloumns Table for Values
                    oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 10)
                    oTable.Range.Style = "No Spacing"
                    oTable.Range.Font.Size = 9
                    'oTable.Range.Font.Bold = True
                    'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
                    'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
                    'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle

                    PWNRPT_Headings(oTable, oWord, "Status: " & sType)

                    r = 4

                End If

                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = reader(1).ToString()
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 2).Range.Text = reader(2).ToString()
                oTable.Cell(r, 2).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 3).Range.Text = reader(3).ToString()
                oTable.Cell(r, 3).Range.Font.Bold = False
                oTable.Cell(r, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 4).Range.Text = reader(4).ToString()
                oTable.Cell(r, 4).Range.Font.Bold = False
                oTable.Cell(r, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 5).Range.Text = reader(5).ToString()
                oTable.Cell(r, 5).Range.Font.Bold = False
                oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 6).Range.Text = reader(6).ToString()
                oTable.Cell(r, 6).Range.Font.Bold = False
                oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 7).Range.Text = reader(7).ToString()
                oTable.Cell(r, 7).Range.Font.Bold = False
                oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 8).Range.Text = GetValue(reader(8).ToString())
                oTable.Cell(r, 8).Range.Font.Bold = False
                oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 9).Range.Text = GetValue(reader(9).ToString())
                lGroupTotal = lGroupTotal + reader(9).ToString()
                oTable.Cell(r, 9).Range.Font.Bold = False
                oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                r = r + 1
            End While
            reader.Close()

            oTable.Rows.Add()
            oTable.Cell(r, 1).Range.Text = ""
            oTable.Cell(r, 2).Range.Text = ""
            oTable.Cell(r, 3).Range.Text = ""
            oTable.Cell(r, 4).Range.Text = ""
            oTable.Cell(r, 5).Range.Text = ""
            oTable.Cell(r, 6).Range.Text = ""
            oTable.Cell(r, 7).Range.Text = ""
            oTable.Cell(r, 8).Range.Text = ""
            oTable.Cell(r, 9).Range.Text = ""
            oTable.Cell(r, 7).Merge(oTable.Cell(r, 8))

            oTable.Cell(r, 7).Range.Text = "Status Total:"
            oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(r, 7).Range.Font.Bold = True
            oTable.Cell(r, 8).Range.Text = GetValue(lGroupTotal)
            oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(r, 8).Range.Font.Bold = True

            'lBigTotal = lBigTotal + lGroupTotal
            'r = r + 1
            'oTable.Rows.Add()
            'oTable.Cell(r, 1).Range.Text = ""
            'oTable.Cell(r, 2).Range.Text = ""
            'oTable.Cell(r, 3).Range.Text = ""
            'oTable.Cell(r, 4).Range.Text = ""
            'oTable.Cell(r, 5).Range.Text = ""
            'oTable.Cell(r, 6).Range.Text = ""
            'oTable.Cell(r, 7).Range.Text = ""
            'oTable.Cell(r, 8).Range.Text = ""
            'r = r + 1
            'oTable.Rows.Add()
            'oTable.Cell(r, 1).Range.Text = ""
            'oTable.Cell(r, 2).Range.Text = ""
            'oTable.Cell(r, 3).Range.Text = ""
            'oTable.Cell(r, 4).Range.Text = ""
            'oTable.Cell(r, 5).Range.Text = ""
            'oTable.Cell(r, 6).Range.Text = ""
            'oTable.Cell(r, 7).Range.Text = ""
            'oTable.Cell(r, 8).Range.Text = ""
            'oTable.Cell(r, 4).Merge(oTable.Cell(r, 7))

            'oTable.Cell(r, 4).Range.Text = "Total Value of Pawn Stock:"
            'oTable.Cell(r, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oTable.Cell(r, 4).Range.Font.Bold = True
            'oTable.Cell(r, 5).Range.Text = GetValue(lBigTotal)
            'oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            'oTable.Cell(r, 5).Range.Font.Bold = True

            oTable.Cell(3, 4).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 4).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 5).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 6).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(3, 7).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle

            If bVlauesFound <> True Then
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 2))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

    Public Sub PWNRPT_Headings(otable As Object, oword As Object, Heading As String)

        otable.Cell(1, 1).Range.Text = Heading
        otable.Columns.Item(1).Width = oword.CentimetersToPoints(1.25)
        otable.Cell(1, 1).Range.Font.Bold = True
        otable.Cell(1, 2).Range.Text = ""
        otable.Columns.Item(2).Width = oword.CentimetersToPoints(5.5)
        otable.Cell(1, 2).Range.Font.Bold = True
        otable.Cell(1, 3).Range.Text = ""
        'oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(2)
        otable.Cell(1, 3).Range.Font.Bold = True
        otable.Cell(1, 4).Range.Text = ""
        otable.Columns.Item(4).Width = oword.CentimetersToPoints(3.2)
        otable.Cell(1, 4).Range.Font.Bold = True
        otable.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        otable.Cell(1, 5).Range.Text = ""
        otable.Columns.Item(5).Width = oword.CentimetersToPoints(1.4)
        otable.Cell(1, 5).Range.Font.Bold = True
        otable.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(1, 6).Range.Text = ""
        otable.Columns.Item(6).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(1, 6).Range.Font.Bold = True
        otable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(1, 7).Range.Text = ""
        otable.Columns.Item(7).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(1, 7).Range.Font.Bold = True
        otable.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(1, 8).Range.Text = ""
        'oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(2)
        otable.Cell(1, 8).Range.Font.Bold = True
        otable.Cell(1, 9).Range.Text = ""
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        otable.Cell(1, 9).Range.Font.Bold = True

        otable.Rows.Add()

        otable.Cell(2, 1).Range.Text = ""
        otable.Columns.Item(1).Width = oword.CentimetersToPoints(1.25)
        otable.Cell(2, 1).Range.Font.Bold = True
        otable.Cell(2, 2).Range.Text = ""
        otable.Columns.Item(2).Width = oword.CentimetersToPoints(5.5)
        otable.Cell(2, 2).Range.Font.Bold = True
        otable.Cell(2, 3).Range.Text = ""
        'oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(2)
        otable.Cell(2, 3).Range.Font.Bold = True
        otable.Cell(2, 4).Range.Text = "Quantities:"
        otable.Columns.Item(4).Width = oword.CentimetersToPoints(3.2)
        otable.Cell(2, 4).Range.Font.Bold = True
        otable.Cell(2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 5).Range.Text = ""
        otable.Columns.Item(5).Width = oword.CentimetersToPoints(1.4)
        otable.Cell(2, 5).Range.Font.Bold = True
        otable.Cell(2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 6).Range.Text = ""
        otable.Columns.Item(6).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(2, 6).Range.Font.Bold = True
        otable.Cell(2, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 7).Range.Text = ""
        otable.Columns.Item(7).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(2, 7).Range.Font.Bold = True
        otable.Cell(2, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 8).Range.Text = ""
        'oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(2)
        otable.Cell(2, 8).Range.Font.Bold = True
        otable.Cell(2, 9).Range.Text = ""
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        otable.Cell(2, 9).Range.Font.Bold = True

        otable.Rows.Add()

        otable.Cell(3, 1).Range.Text = "PS ID:"
        otable.Columns.Item(1).Width = oword.CentimetersToPoints(1.25)
        otable.Cell(3, 1).Range.Font.Bold = True
        otable.Cell(3, 2).Range.Text = "Description:"
        otable.Columns.Item(2).Width = oword.CentimetersToPoints(5.5)
        otable.Cell(3, 2).Range.Font.Bold = True
        otable.Cell(3, 3).Range.Text = "Serial Nr:"
        otable.Columns.Item(3).Width = oword.CentimetersToPoints(3.2)
        otable.Cell(3, 3).Range.Font.Bold = True
        otable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(3, 4).Range.Text = "Total:"
        otable.Columns.Item(4).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(3, 4).Range.Font.Bold = True
        otable.Cell(3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(3, 5).Range.Text = "LayBy:"
        otable.Columns.Item(5).Width = oword.CentimetersToPoints(1.4)
        otable.Cell(3, 5).Range.Font.Bold = True
        otable.Cell(3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(3, 6).Range.Text = "Sold:"
        otable.Columns.Item(6).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(3, 6).Range.Font.Bold = True
        otable.Cell(3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(3, 7).Range.Text = "Avail.:"
        otable.Columns.Item(7).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(3, 7).Range.Font.Bold = True
        otable.Cell(3, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(3, 8).Range.Text = "UnitPrice:"
        'oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(2)
        otable.Cell(3, 8).Range.Font.Bold = True
        otable.Cell(3, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(3, 9).Range.Text = "Realizable Income:"
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        otable.Cell(3, 9).Range.Font.Bold = True
        otable.Cell(3, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

        otable.Cell(1, 1).Merge(otable.Cell(1, 3))
        otable.Cell(1, 5).Merge(otable.Cell(1, 6))
        otable.Cell(2, 4).Merge(otable.Cell(2, 7))

    End Sub

    Public Sub REALINC_Headings(otable As Object, oword As Object)
        otable.Cell(1, 1).Range.Text = ""
        otable.Columns.Item(1).Width = oword.CentimetersToPoints(1.25)
        otable.Cell(1, 1).Range.Font.Bold = True
        otable.Cell(1, 2).Range.Text = ""
        otable.Columns.Item(2).Width = oword.CentimetersToPoints(5.5)
        otable.Cell(1, 2).Range.Font.Bold = True
        otable.Cell(1, 3).Range.Text = ""
        'oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(2)
        otable.Cell(1, 3).Range.Font.Bold = True
        otable.Cell(1, 4).Range.Text = ""
        otable.Columns.Item(4).Width = oword.CentimetersToPoints(3.2)
        otable.Cell(1, 4).Range.Font.Bold = True
        otable.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(1, 5).Range.Text = "Quantities:"
        otable.Columns.Item(5).Width = oword.CentimetersToPoints(1.4)
        otable.Cell(1, 5).Range.Font.Bold = True
        otable.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(1, 6).Range.Text = ""
        otable.Columns.Item(6).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(1, 6).Range.Font.Bold = True
        otable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(1, 7).Range.Text = ""
        otable.Columns.Item(7).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(1, 7).Range.Font.Bold = True
        otable.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(1, 8).Range.Text = ""
        'oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(2)
        otable.Cell(1, 8).Range.Font.Bold = True
        otable.Cell(1, 9).Range.Text = ""
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        otable.Cell(1, 9).Range.Font.Bold = True

        otable.Rows.Add()

        otable.Cell(2, 1).Range.Text = "PS ID:"
        otable.Columns.Item(1).Width = oword.CentimetersToPoints(1.25)
        otable.Cell(2, 1).Range.Font.Bold = True
        otable.Cell(2, 2).Range.Text = "Description:"
        otable.Columns.Item(2).Width = oword.CentimetersToPoints(5.5)
        otable.Cell(2, 2).Range.Font.Bold = True
        otable.Cell(2, 3).Range.Text = "Serial Nr:"
        otable.Columns.Item(3).Width = oword.CentimetersToPoints(3.2)
        otable.Cell(2, 3).Range.Font.Bold = True
        otable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 4).Range.Text = "PnCID:"
        otable.Columns.Item(4).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(2, 4).Range.Font.Bold = True
        otable.Cell(2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 5).Range.Text = "Total:"
        otable.Columns.Item(5).Width = oword.CentimetersToPoints(1.4)
        otable.Cell(2, 5).Range.Font.Bold = True
        otable.Cell(2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 6).Range.Text = "Sold:"
        otable.Columns.Item(6).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(2, 6).Range.Font.Bold = True
        otable.Cell(2, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 7).Range.Text = "Avail.:"
        otable.Columns.Item(7).Width = oword.CentimetersToPoints(1.3)
        otable.Cell(2, 7).Range.Font.Bold = True
        otable.Cell(2, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 8).Range.Text = "UnitPrice:"
        'oTable.Columns.Item(8).Width = oWord.CentimetersToPoints(2)
        otable.Cell(2, 8).Range.Font.Bold = True
        otable.Cell(2, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable.Cell(2, 9).Range.Text = "Realizable Income:"
        'oTable.Columns.Item(9).Width = oWord.CentimetersToPoints(2)
        otable.Cell(2, 9).Range.Font.Bold = True
        otable.Cell(2, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

        otable.Cell(1, 5).Merge(otable.Cell(1, 7))
    End Sub

    Private Sub cmdRealIncome_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles cmdRealIncome.Click


        Dim bVlauesFound As Boolean

        Dim r As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph

        Dim lGroupTotal As Long
        Dim lBigTotal As Long
        Dim sType As String
        Dim sPNCID As String

        sType = ""
        lGroupTotal = 0
        lBigTotal = 0
        sPNCID = ""
        bVlauesFound = False

        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Report on Realizable Income"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 4 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9

            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " "
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            lGroupTotal = 0

            'Add some text after the table.
            'oTable.Range.InsertParagraphAfter()
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Range.InsertParagraphBefore()
            oPara2.Range.Style = "No Spacing"
            oPara2.Range.Font.Size = 9
            oPara2.Range.Font.Bold = True

            'Insert x Rows x 10 Coloumns Table for Values
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 10)
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            'oTable.Range.Font.Bold = True
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle

            REALINC_Headings(oTable, oWord)

            r = 3

            'sSQL = "SELECT a.Status, a.PSID, a.Description, a.SerialNr, a.PnCID as PnCID, a.Quantity as QuantityTotal, a.QuantitySold, (a.Quantity - a.QuantityLayBuy - a.QuantitySold) as [Quantity Available], b.PurchaseAmount as UnitPrice, b.BuyBackAmount as [Realizable Profit] FROM PawnStock a RIGHT JOIN PawningTransactions b on b.PnCID = a.PnCID where a.status=1 ORDER BY PSID"
            'sSQL = "SELECT a.Status, a.PSID, a.Description, a.SerialNr, a.PnCID as PnCID, a.Quantity as QtyTotal, a.QuantitySold, (a.Quantity - a.QuantityLayBuy - a.QuantitySold) as QtyAvail, b.PurchaseAmount as UnitPrice, (QtyAvail * b.BuyBackAmount) as RealProfit FROM PawnStock a RIGHT JOIN PawningTransactions b on b.PnCID = a.PnCID where a.status=1 ORDER BY PSID"
            sSQL = "SELECT a.Status, a.PSID, a.Description, a.SerialNr, a.PnCID as PnCID, a.Quantity as QtyTotal, a.QuantitySold, (a.Quantity - a.QuantityLayBuy - a.QuantitySold) as QtyAvail, b.PurchaseAmount as UnitPrice, b.BuyBackAmount as RealProfit FROM PawnStock a RIGHT JOIN PawningTransactions b on b.PnCID = a.PnCID where a.status=1 ORDER BY PSID"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = reader(1).ToString()
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 2).Range.Text = reader(2).ToString()
                oTable.Cell(r, 2).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 3).Range.Text = reader(3).ToString()
                oTable.Cell(r, 3).Range.Font.Bold = False
                oTable.Cell(r, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 4).Range.Text = reader(4).ToString()
                oTable.Cell(r, 4).Range.Font.Bold = False
                oTable.Cell(r, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 5).Range.Text = reader(5).ToString()
                oTable.Cell(r, 5).Range.Font.Bold = False
                oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 6).Range.Text = reader(6).ToString()
                oTable.Cell(r, 6).Range.Font.Bold = False
                oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 7).Range.Text = reader(7).ToString()
                oTable.Cell(r, 7).Range.Font.Bold = False
                oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                If sPNCID = reader(4).ToString Then
                    oTable.Cell(r, 8).Range.Text = "-"
                    oTable.Cell(r, 8).Range.Font.Bold = False
                    oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    oTable.Cell(r, 9).Range.Text = "-"
                    oTable.Cell(r, 9).Range.Font.Bold = False
                    oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    sPNCID = reader(4).ToString
                Else
                    oTable.Cell(r, 8).Range.Text = GetValue(reader(8).ToString())
                    oTable.Cell(r, 8).Range.Font.Bold = False
                    oTable.Cell(r, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    oTable.Cell(r, 9).Range.Text = GetValue(reader(9).ToString())
                    lGroupTotal = lGroupTotal + reader(9).ToString()
                    oTable.Cell(r, 9).Range.Font.Bold = False
                    oTable.Cell(r, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    sPNCID = reader(4).ToString
                End If
                r = r + 1

            End While
            reader.Close()

            oTable.Rows.Add()
            oTable.Cell(r, 1).Range.Text = ""
            oTable.Cell(r, 2).Range.Text = ""
            oTable.Cell(r, 3).Range.Text = ""
            oTable.Cell(r, 4).Range.Text = ""
            oTable.Cell(r, 5).Range.Text = ""
            oTable.Cell(r, 6).Range.Text = ""
            oTable.Cell(r, 7).Range.Text = ""
            oTable.Cell(r, 8).Range.Text = ""
            oTable.Cell(r, 9).Range.Text = ""

            oTable.Cell(r, 4).Merge(oTable.Cell(r, 7))
            oTable.Cell(r, 4).Range.Text = "Realizable Income Total:"
            oTable.Cell(r, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(r, 4).Range.Font.Bold = True

            oTable.Cell(r, 5).Merge(oTable.Cell(r, 6))
            oTable.Cell(r, 5).Range.Text = GetValue(lGroupTotal)
            oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(r, 5).Range.Font.Bold = True

            oTable.Cell(2, 5).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(2, 5).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(2, 6).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(2, 6).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(2, 7).Borders(WdBorderType.wdBorderTop).Color = Word.WdColor.wdColorBlack
            oTable.Cell(2, 7).Borders(WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle

            If bVlauesFound <> True Then
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 2))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

    Private Sub cmdStockSale_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStockSale.Click

        Dim iRows As Integer
        Dim r As Integer, c As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        'Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
        'Dim oRng As Word.Range
        'Dim Pos As Double


        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Cash & BuyBack: Stock Sale"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 2 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "" '"Report Creation: " & vbNewLine & Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " " & vbNewLine & "Tax Invoice"
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            'Add some text after the table.
            'oTable.Range.InsertParagraphAfter()
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Range.InsertParagraphBefore()
            oPara2.Range.Style = "No Spacing"
            'oPara2.Range.Text = "And here's another table:"

            'Insert x Rows x 7 Coloumns Table for Values
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 7)
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            oTable.Cell(1, 1).Range.Text = "Date:"
            oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(2)
            oTable.Cell(1, 1).Range.Font.Bold = True
            oTable.Cell(1, 2).Range.Text = "ID:"
            oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(2.25)
            oTable.Cell(1, 2).Range.Font.Bold = True
            oTable.Cell(1, 3).Range.Text = "Qty:"
            oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(1)
            oTable.Cell(1, 3).Range.Font.Bold = True
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 4).Range.Text = "Description:"
            oTable.Columns.Item(4).Width = oWord.CentimetersToPoints(6)
            oTable.Cell(1, 4).Range.Font.Bold = True
            oTable.Cell(1, 5).Range.Text = "SerialNr:"
            oTable.Columns.Item(5).Width = oWord.CentimetersToPoints(3)
            oTable.Cell(1, 5).Range.Font.Bold = True
            oTable.Cell(1, 6).Range.Text = "Price:"
            oTable.Columns.Item(6).Width = oWord.CentimetersToPoints(2)
            oTable.Cell(1, 6).Range.Font.Bold = True
            oTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(1, 7).Range.Text = "SubTotal:"
            oTable.Columns.Item(7).Width = oWord.CentimetersToPoints(2.25)
            oTable.Cell(1, 7).Range.Font.Bold = True
            oTable.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

            For c = 1 To 7
                oTable.Cell(2, c).Range.Text = ""
            Next

            r = 3
            For iRows = 0 To iSaleItemCount
                'Public arrItems(9, 7)  'ItemId, Quatity, Description, Serialnr, Price, SubTotal, SelDate
                oTable.Cell(r, 1).Range.Text = arrItems(iRows, 6)
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.Text = arrItems(iRows, 0)
                oTable.Cell(r, 2).Range.Font.Bold = False
                oTable.Cell(r, 3).Range.Text = arrItems(iRows, 1)
                oTable.Cell(r, 3).Range.Font.Bold = False
                oTable.Cell(r, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 4).Range.Text = arrItems(iRows, 2)
                oTable.Cell(r, 4).Range.Font.Bold = False
                oTable.Cell(r, 5).Range.Text = arrItems(iRows, 3)
                oTable.Cell(r, 5).Range.Font.Bold = False
                oTable.Cell(r, 6).Range.Text = arrItems(iRows, 4)
                oTable.Cell(r, 6).Range.Font.Bold = False
                oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 7).Range.Text = arrItems(iRows, 5)
                oTable.Cell(r, 7).Range.Font.Bold = False
                oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                r = r + 1
                oTable.Rows.Add()
            Next iRows

            'oTable.Rows.Add()
            oTable.Cell(r, 1).Range.Text = "*Items sold 'as is'"
            oTable.Cell(r, 1).Range.Font.Bold = False
            oTable.Cell(r, 1).Merge(oTable.Cell(r, 2))

            r = r + 1
            oTable.Rows.Add()
            oTable.Cell(r, 1).Range.Text = "*Vat incl."
            oTable.Cell(r, 1).Range.Font.Bold = False
            oTable.Cell(r, 4).Range.Text = "Grand Total:"
            oTable.Cell(r, 4).Range.Font.Bold = True
            oTable.Cell(r, 4).Merge(oTable.Cell(r, 5))
            oTable.Cell(r, 5).Range.Text = sSaleTotal
            oTable.Cell(r, 5).Range.Font.Bold = True

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try
    End Sub

    Private Sub cmdMoveExpiredStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPawnReport.Click
        Dim bVlauesFound As Boolean

        Dim r As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph

        bVlauesFound = False

        Try
            'Start Word and open the document template.
            oWord = CreateObject("Word.Application")
            'oWord.Visible = True
            oDoc = oWord.Documents.Add
            oDoc.Range.NoProofing = True
            'oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.75)
            'oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.75)
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5)
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5)
            oDoc.Sections(1).Footers(1).PageNumbers.Add()

            'Insert Heading at the beginning of the document.
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "Expired Pawn"
            'oPara1.Range.Style = "Heading 1"
            oPara1.Range.Font.Bold = True
            oPara1.Range.Font.Size = 26
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'oPara1.Range.Font.Underline = True
            oPara1.Format.SpaceAfter = 10    '24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter()

            'Insert a 4 Rows x 3 Coloumns table, for Address
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            'oTable.Range.ParagraphFormat.SpaceAfter = 6
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9

            oTable.Cell(1, 1).Range.Text = p_AddressWH & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWH
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 2).Range.Text = "Report Creation: " & vbNewLine & _
                                            Format(Now, "dd MMMM yyyy HH:mm:ss")
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 3).Range.Text = p_AddressWB & vbNewLine & _
                                           " " & vbNewLine & _
                                           " " & vbNewLine & _
                                           p_TelWB
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(2, 1).Merge(oTable.Cell(2, 3))
            oTable.Cell(2, 1).Range.Text = " "
            oTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(3, 1).Merge(oTable.Cell(3, 3))
            oTable.Cell(3, 1).Range.Text = "The following items must be moved from pawn stock to the general sales area and marked with PSID and Price"
            oTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(3, 1).Range.Bold = True
            oTable.Cell(4, 1).Merge(oTable.Cell(4, 3))
            oTable.Cell(4, 1).Range.Text = vbTab & " "
            oTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Range.InsertParagraphBefore()
            oPara2.Range.Style = "No Spacing"
            'oPara2.Range.Text = "And here's another table:"

            'Insert x Rows x 6 Coloumns Table for Values
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 2, 6)
            oTable.Range.Style = "No Spacing"
            oTable.Range.Font.Size = 9
            'oTable.Range.Font.Bold = True
            'oTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            'oTable.Borders.InsideColor = Word.WdColor.wdColorBlack
            'oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            oTable.Cell(1, 1).Range.Text = "CBBID:"
            oTable.Columns.Item(1).Width = oWord.CentimetersToPoints(1.3)
            oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 1).Range.Font.Bold = True
            oTable.Cell(1, 2).Range.Text = "Qty:"
            oTable.Columns.Item(2).Width = oWord.CentimetersToPoints(1.25)
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(1, 2).Range.Font.Bold = True
            oTable.Cell(1, 3).Range.Text = "Description:"
            oTable.Columns.Item(3).Width = oWord.CentimetersToPoints(7)
            oTable.Cell(1, 3).Range.Font.Bold = True
            oTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 4).Range.Text = "SerialNr:"
            oTable.Columns.Item(4).Width = oWord.CentimetersToPoints(3)
            oTable.Cell(1, 4).Range.Font.Bold = True
            oTable.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oTable.Cell(1, 5).Range.Text = "Price:"
            oTable.Columns.Item(5).Width = oWord.CentimetersToPoints(2.75)
            oTable.Cell(1, 5).Range.Font.Bold = True
            oTable.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            oTable.Cell(1, 6).Range.Text = "PS ID:"
            oTable.Columns.Item(6).Width = oWord.CentimetersToPoints(1.5)
            oTable.Cell(1, 6).Range.Font.Bold = True
            oTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

            For c = 1 To 6
                oTable.Cell(2, c).Range.Text = ""
            Next

            r = 3
            sSQL = "SELECT PawnStock.PnCID, PawnStock.Quantity, PawnStock.Description, PawnStock.SerialNr, PawnStock.Price, PawnStock.PSID FROM PawningTransactions INNER JOIN PawnStock ON PawningTransactions.PnCID = PawnStock.PnCID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <=  #" & FormatDateTime(Now, DateFormat.ShortDate) & "#) AND (PawningTransactions.Status = 1)"
            com.CommandText = sSQL
            Dim reader As OleDbDataReader = com.ExecuteReader()
            While reader.Read()
                Me.PictureBox1.Refresh()
                bVlauesFound = True
                'Console.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())
                'Debug.WriteLine(reader(0).ToString() + reader(1).ToString() + reader(2).ToString())

                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = reader(0).ToString()
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 2).Range.Text = reader(1).ToString()
                oTable.Cell(r, 2).Range.Font.Bold = False
                oTable.Cell(r, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, 3).Range.Text = reader(2).ToString()
                oTable.Cell(r, 3).Range.Font.Bold = False
                oTable.Cell(r, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 4).Range.Text = reader(3).ToString()
                oTable.Cell(r, 4).Range.Font.Bold = False
                oTable.Cell(r, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                oTable.Cell(r, 5).Range.Text = GetValue(reader(4).ToString())
                oTable.Cell(r, 5).Range.Font.Bold = False
                oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, 6).Range.Text = reader(5).ToString()
                oTable.Cell(r, 6).Range.Font.Bold = False
                oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                r = r + 1
            End While
            reader.Close()

            If bVlauesFound <> True Then
                oTable.Rows.Add()
                oTable.Cell(r, 1).Range.Text = "No Values Found."
                oTable.Cell(r, 1).Range.Font.Bold = False
                oTable.Cell(r, 1).Merge(oTable.Cell(r, 2))
            End If

            oWord.Visible = True

            'All done. Close this form.
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ErrorToFile(Me.GetType().Name.ToString, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString, ex.Message, ex.StackTrace)
            Me.Close()
        End Try

    End Sub

End Class
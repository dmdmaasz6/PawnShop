Imports System.Globalization
Imports System.Threading

Public Class frmMain

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        WriteRegistryDate(Now)
        CloseApp()

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim newCulture As CultureInfo = DirectCast(System.Threading.Thread.CurrentThread.CurrentCulture.Clone(), CultureInfo)
        newCulture.DateTimeFormat.ShortDatePattern = "MM/dd/yyyy"
        newCulture.DateTimeFormat.DateSeparator = "/"
        Thread.CurrentThread.CurrentCulture = newCulture

        dtPawnFrom.Value = DateAdd(DateInterval.Month, -1, Now)

        ReadRegValues()
        If p_DBLocation <> "" Then
            DBConnect(p_DBLocation)
        End If

        My.Settings("C_bbConnectionString") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & p_DBLocation

        p_Currency = System.Globalization.RegionInfo.CurrentRegion.CurrencySymbol
        p_Search_NSID = ""
        p_Search_PSID = ""

        sTemplatePath = Application.StartupPath & "\templates\"
        sTemplateShop = p_Branch

        SetAddresses()

        SS1.Text = SS1.Text & Replace(p_CashRegister, "&", "&&")
        SS2.Text = SS2.Text & UCase(p_Branch)
        SS3.Text = TimeOfDay

        Timer1.Enabled = True

        CheckDate()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        SS3.Text = TimeOfDay
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        CloseApp()

    End Sub


    Private Sub ReadRegValues()
        Dim bReadValues As Boolean
        bReadValues = False

        If ReadRegistryBrnch() = "" Then bReadValues = True
        If ReadRegistryCR() = "" Then bReadValues = True
        If ReadRegistryDB() = "" Then bReadValues = True

        If ReadRegistryAgreePath() = "" Then
            p_AgreementPath = "C:\Agreements"
        End If

        ReadRegistryDate()
        ReadRegistryAudit()

        If bReadValues = True Then
            MsgBox("Not all the values are properly set" & vbCrLf & "This must be corrected before the application can be used!", vbCritical)
            OpenConfig()
        End If

        'If ReadRegistryDate() = "" Then
        '    '   MsgBox("A")
        '    'Else
        '    WriteRegistryDate(FormatDateTime(Now, DateFormat.ShortDate))
        'End If

    End Sub

    Private Sub CloseApp()
        DBClose()
        End

    End Sub

    Private Sub DBLoactionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DBLoactionToolStripMenuItem.Click
        frmSettings.ShowDialog(Me)

    End Sub

    Private Sub CashRegisterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CashRegisterToolStripMenuItem.Click
        OpenConfig()

    End Sub

    Private Sub OpenConfig()
        Dim sPswd As String

        sPswd = InputBox("Password is required to change settings!" & vbCrLf & "Please specify...", "Setting Password.")
        If UCase(sPswd) = "7777" Then
            frmConfig.ShowDialog(Me)
        Else
            MsgBox("Illegal password", vbInformation, "Password incorrect")
        End If

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.cmdOK.Visible = True
        frmAbout.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D
        frmAbout.StartPosition = FormStartPosition.CenterParent
        frmAbout.ShowDialog(Me)

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles dtReportFrom.ValueChanged

    End Sub

    Private Sub cmdCustAdd_Click(sender As Object, e As EventArgs) Handles cmdCustAdd.Click
        With frmCustMngmnt
            .txtCustID.Enabled = False
            .cmdUpdateCustomerInformation.Enabled = False
            .cmdAddCustomer.Enabled = True
            .cmdUseThisCustomer.Visible = False
            .ShowDialog(Me)
        End With


    End Sub

    Private Sub cmdReportDaily_Click(sender As Object, e As EventArgs) Handles cmdReportDaily.Click
        With frmReports
            .sDateTo = FormatDateTime(dtReportTo.Value, DateFormat.ShortDate)
            .sDateFrom = ""
            .Label1.Text = "Generating Daily Report for " & vbCrLf & FormatDateTime(dtReportTo.Value, DateFormat.LongDate) & " ..."
            .Width = 485 + 70 + frmReports.PictureBox1.Width
            .Height = 67
            .Report = frmReports.ReportType.DailyReport
            .ShowDialog(Me)
        End With

    End Sub

    Private Sub cmdCustsearch_Click(sender As Object, e As EventArgs) Handles cmdCustsearch.Click
        With frmCustMngmnt
            .cmdAddCustomer.Enabled = False
            .cmdUpdateCustomerInformation.Enabled = False
            .cmdShowHistory.Visible = True
            .cmdShowHistory.Location = New Point(205, 508)
            .cmdUseThisCustomer.Visible = False
            .ShowDialog(Me)
        End With
    End Sub

    Private Sub cmdCustUpdate_Click(sender As Object, e As EventArgs) Handles cmdCustUpdate.Click
        With frmCustMngmnt
            .cmdAddCustomer.Enabled = False
            .cmdUpdateCustomerInformation.Enabled = True
            .cmdUseThisCustomer.Visible = False
            .ShowDialog(Me)
        End With

    End Sub

    Private Sub cmdReportCustom_Click(sender As Object, e As EventArgs) Handles cmdReportCustom.Click
        If DateDiff(DateInterval.Day, dtReportTo.Value, dtReportFrom.Value) > 0 Then
            MsgBox("Start Date cannot be greater than End Date!", MsgBoxStyle.Exclamation + vbOKOnly, "Cash & BuyBack Error")
        Else
            With frmReports
                .sDateTo = FormatDateTime(dtReportTo.Value, DateFormat.ShortDate)
                .sDateFrom = FormatDateTime(dtReportFrom.Value, DateFormat.ShortDate)
                .Label1.Text = "Generating Custom Report for the period: " & vbCrLf & FormatDateTime(dtReportFrom.Value, DateFormat.LongDate) & " - " & FormatDateTime(dtReportTo.Value, DateFormat.LongDate) & " ..."
                .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
                .Height = 67
                .Report = frmReports.ReportType.CustomReport
                .ShowDialog(Me)
            End With
        End If
    End Sub

    Private Sub cmdPawnHistory_Click(sender As Object, e As EventArgs) Handles cmdPawnHistory.Click
        Dim tStart
        Dim tEnd
        If DateDiff(DateInterval.Day, dtPawnTo.Value, dtPawnFrom.Value) > 0 Then
            MsgBox("Start Date cannot be greater than End Date!", MsgBoxStyle.Exclamation + vbOKOnly, "Cash & BuyBack Error")
        Else
            tStart = FormatDateTime(Now, DateFormat.LongTime)

            With frmReports
                .sDateTo = FormatDateTime(dtPawnTo.Value, DateFormat.ShortDate)
                .sDateFrom = FormatDateTime(dtPawnFrom.Value, DateFormat.ShortDate)
                .Label1.Text = "Generating Pawn History for the period: " & vbCrLf & FormatDateTime(dtPawnFrom.Value, DateFormat.LongDate) & " - " & FormatDateTime(dtPawnTo.Value, DateFormat.LongDate) & " ..."
                .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
                .Height = 67
                .Report = frmReports.ReportType.PawnHisotry
                .ShowDialog(Me)
            End With


            'With frmRptPawnHistory
            '    .sDateTo = FormatDateTime(dtPawnTo.Value, DateFormat.ShortDate)
            '    .sDateFrom = FormatDateTime(dtPawnFrom.Value, DateFormat.ShortDate)
            '    '                .Label1.Text = "Generating Pawn History for the period: " & FormatDateTime(dtPawnFrom.Value, DateFormat.LongDate) & " - " & FormatDateTime(dtPawnTo.Value, DateFormat.LongDate) & " ..."
            '    '               .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
            '    '                .Height = 67
            '    '               .Report = frmReports.ReportType.PawnHisotry
            '    .ShowDialog(Me)
            'End With

            tEnd = FormatDateTime(Now, DateFormat.LongTime)
            MsgBox("Start Time: " & tStart.ToString & vbCrLf _
                   & "End Time: " & tEnd.ToString & vbCrLf _
                   & "Time Taken: " & DateDiff(DateInterval.Second, tStart, tEnd))

        End If
    End Sub

    Private Sub cmdStockSearch_Click(sender As Object, e As EventArgs) Handles cmdStockSearch.Click
        frmStockSearch.cmdSell.Visible = False
        frmStockSearch.ShowDialog(Me)

    End Sub

    Private Sub cmdStockCdStnd_Click(sender As Object, e As EventArgs) Handles cmdStockCdStnd.Click
        'With frmReports
        '    .Label1.Text = "Generating Standard Stock Code List ..."
        '    .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
        '    .Height = 67
        '    .Report = frmReports.ReportType.StdStockList
        '    .ShowDialog(Me)
        'End With
        Me.Cursor = Cursors.WaitCursor
        frmRptStdStockCode.ShowDialog(Me)
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cmdStockCdTango_Click(sender As Object, e As EventArgs) Handles cmdStockCdTango.Click
        'With frmReports
        '    .Label1.Text = "Generating Tango Stock Code List ..."
        '    .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
        '    .Height = 67
        '    .Report = frmReports.ReportType.TangoStockList
        '    .ShowDialog(Me)
        'End With
        Me.Cursor = Cursors.WaitCursor
        frmRptTangoStockCode.ShowDialog(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cmdStockReport_Click(sender As Object, e As EventArgs) Handles cmdStockReport.Click
        'With frmReports
        '    .Label1.Text = "Generating Stock Report ..."
        '    .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
        '    .Height = 67
        '    .Report = frmReports.ReportType.StockReport
        '    .ShowDialog(Me)
        'End With
        Me.Cursor = Cursors.WaitCursor
        frmRptNormalStock.ShowDialog(Me)
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cmdStackPawnRpt_Click(sender As Object, e As EventArgs) Handles cmdStackPawnRpt.Click
        'With frmReports
        '    .Label1.Text = "Generating Pawn Report ..."
        '    .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
        '    .Height = 67
        '    .Report = frmReports.ReportType.PawnReport
        '    .ShowDialog(Me)
        'End With
        Me.Cursor = Cursors.WaitCursor
        frmRptPawnStock.ShowDialog(Me)
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cmdStockTango_Click(sender As Object, e As EventArgs) Handles cmdStockTango.Click
        frmTangoStockManagment.ShowDialog(Me)

    End Sub

    Private Sub cmdStockStnd_Click(sender As Object, e As EventArgs) Handles cmdStockStnd.Click
        frmStdStockManagment.ShowDialog(Me)

    End Sub

    Private Sub cmdStockIncome_Click(sender As Object, e As EventArgs) Handles cmdStockIncome.Click
        'With frmReports
        '    .Label1.Text = "Generating Realizable Income Report ..."
        '    .Width = frmReports.Label1.Width + 70 + frmReports.PictureBox1.Width
        '    .Height = 67
        '    .Report = frmReports.ReportType.RealIncome
        '    .ShowDialog(Me)
        'End With
        Me.Cursor = Cursors.WaitCursor
        frmRptRealIncome.ShowDialog(Me)
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cmdCashOther_Click(sender As Object, e As EventArgs) Handles cmdCashOther.Click
        frmOtherExpenses.ShowDialog(Me)

    End Sub

    Private Sub cmdCashStock_Click(sender As Object, e As EventArgs) Handles cmdCashStock.Click
        frmStockPurchase.Width = 489
        frmStockPurchase.ShowDialog(Me)

    End Sub

    Private Sub cmdCashSale_Click(sender As Object, e As EventArgs) Handles cmdCashSale.Click
        frmStockSale.ShowDialog(Me)

    End Sub

    Private Sub cmdPawnOut_Click(sender As Object, e As EventArgs) Handles cmdPawnOut.Click
        frmPawnOut.ShowDialog(Me)

    End Sub

    Private Sub cmdPawnIn_Click(sender As Object, e As EventArgs) Handles cmdPawnIn.Click
        frmPawnIn.ShowDialog(Me)

    End Sub

    Private Sub cmdPawnExpired_Click(sender As Object, e As EventArgs) Handles cmdPawnExpired.Click
        frmPawnExpired.ShowDialog(Me)

    End Sub
End Class

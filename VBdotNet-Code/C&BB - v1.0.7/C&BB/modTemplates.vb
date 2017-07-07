Imports Microsoft.Office.Interop

Module modTemplates
    'Public arrStockSold(20, 3) As String
    'Public iStock As Integer


    Dim ObjWord As Object
    Dim oDoc As Object
    Dim oTable As Object

    Public Function ValidTemplates(sTemplate As String) As Boolean
        Dim sMsg As String

        ValidTemplates = False

        If Trim(sTemplateShop) = "" Then
            sMsg = "Start-up Paramter not set:"
            GoTo lblErr
        End If

        If System.IO.Directory.Exists(sTemplatePath) <> True Then
            sMsg = "Template Path " & sTemplatePath & " not found:"
            GoTo lblErr
        End If

        If System.IO.Directory.Exists(sTemplatePath & "\" & sTemplateShop) <> True Then
            sMsg = "Template Path " & sTemplatePath & "\" & sTemplateShop & " not found:"
            GoTo lblErr
        End If

        If System.IO.File.Exists(sTemplatePath & "\" & sTemplateShop & "\" & sTemplate) <> True Then
            sMsg = "Template " & sTemplatePath & "\" & sTemplateShop & "\" & sTemplate & " not found:"
            GoTo lblErr
        End If

        ValidTemplates = True

        Exit Function
lblErr:
        MsgBox(sMsg & vbCrLf & vbCrLf & "Please correct the error!", vbCritical + vbOKOnly, "Template Error!")

    End Function

    Public Sub Template_Agreement(sCBBID As String, _
                                       sMonthNR As String, _
                                       sPTTTime As String, _
                                       sPnCID As String, _
                                       sCustName As String, _
                                       sCustID As String, _
                                       sCustTelW As String, _
                                       sCustTelH As String, _
                                       sCustAddr As String, _
                                       sAmountP As String, _
                                       sAmountBB As String, _
                                       sHandInDate As String, _
                                       sExpDate As String)
        Dim sTemplate As String
        Dim sAgreementPath As String

        Dim ObjWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table

        Try

            sTemplate = "agreement.doc"


            If ValidTemplates(sTemplate) <> True Then
                TemplateError()
                Exit Sub
            Else
                sTemplate = sTemplatePath & sTemplateShop & "\" & sTemplate
            End If

            sAgreementPath = Microsoft.VisualBasic.Year(Now)

            If Microsoft.VisualBasic.Month(Now) < 10 Then
                sAgreementPath = sAgreementPath & "0" & Microsoft.VisualBasic.Month(Now)
            Else
                sAgreementPath = sAgreementPath & Microsoft.VisualBasic.Month(Now)
            End If

            If Microsoft.VisualBasic.Day(Now) < 10 Then
                sAgreementPath = sAgreementPath & "0" & Microsoft.VisualBasic.Day(Now)
            Else
                sAgreementPath = sAgreementPath & Microsoft.VisualBasic.Day(Now)
            End If

            sAgreementPath = p_AgreementPath & "\" & sAgreementPath

            If (Not System.IO.Directory.Exists(sAgreementPath)) Then
                System.IO.Directory.CreateDirectory(sAgreementPath)
            End If

            sAgreementPath = sAgreementPath & "\" & "Agreement_" & sPnCID & "_" & sAgreeTime & ".doc"

            System.IO.File.Copy(sTemplate, sAgreementPath)

            ObjWord = CreateObject("Word.Application")
            ObjWord.Documents.Open(sAgreementPath)

            'ObjWord.Application.Visible = True

            Call ReplaceValues(ObjWord, "<CBBID>", sCBBID)
            Call ReplaceValues(ObjWord, "<MONTHNR>", sMonthNR)
            Call ReplaceValues(ObjWord, "<PTTTIME>", sPTTTime)
            Call ReplaceValues(ObjWord, "<PNCID>", sPnCID)

            Call ReplaceValues(ObjWord, "<CusName>", sCustName)
            Call ReplaceValues(ObjWord, "<CusAdr>", sCustAddr)
            Call ReplaceValues(ObjWord, "<CusID>", sCustID)
            Call ReplaceValues(ObjWord, "<CusTelW>", sCustTelW)
            Call ReplaceValues(ObjWord, "<CusTelH>", sCustTelH)
            'Call ReplaceValues(ObjWord,"<CusAdr>", sCustAddr)
            Call ReplaceValues(ObjWord, "<AmountP>", sAmountP)
            Call ReplaceValues(ObjWord, "<AmountBB>", sAmountBB)
            Call ReplaceValues(ObjWord, "<HandInDat>", sHandInDate)
            Call ReplaceValues(ObjWord, "<ExpDat>", sExpDate)

            'StockTable
            Dim r As Integer, c As Integer
            oDoc = ObjWord.ActiveDocument
            oDoc.Range.NoProofing = True
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("STOCKTABLE").Range, iStock, 3)

            oTable.Borders.InsideLineStyle = 0
            oTable.Borders.OutsideLineStyle = 0
            oTable.Range.ParagraphFormat.SpaceAfter = 0
            oTable.Range.ParagraphFormat.LineSpacing = 12
            oTable.Range.Font.Size = 9
            For r = 1 To iStock
                For c = 1 To 3
                    oTable.Cell(r, c).Range.Text = arrStockSold(r - 1, c - 1) 'Data
                Next
            Next
            oTable.Rows(1).Range.Font.Bold = True
            'Table.Rows(1).range.Font.Italic = True
            oTable.Columns(1).Width = ObjWord.InchesToPoints(1.33)
            oTable.Columns(2).Width = ObjWord.InchesToPoints(3.13)
            oTable.Columns(3).Width = ObjWord.InchesToPoints(1.63)

            ObjWord.Application.Visible = True

            ObjWord.ActiveDocument.Save()

            ObjWord = Nothing

        Catch ex As Exception

            ObjWord = Nothing
            System.Windows.Forms.MessageBox.Show("Error in creating Agreement document." & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub Template_AgreementTearOff(sCBBID As String, _
                                              sCustName As String, _
                                              sAmountBB As String, _
                                              sExpDate As String, _
                                              sPnCID As String)
        Dim sTemplate As String
        Dim sAgreementPathTO As String

        Dim ObjWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table

        Try

            sTemplate = "agreement_2.doc"


            If ValidTemplates(sTemplate) <> True Then
                TemplateError()
                Exit Sub
            Else
                sTemplate = sTemplatePath & sTemplateShop & "\" & sTemplate
            End If

            sAgreementPathTO = Microsoft.VisualBasic.Year(Now)

            If Microsoft.VisualBasic.Month(Now) < 10 Then
                sAgreementPathTO = sAgreementPathTO & "0" & Microsoft.VisualBasic.Month(Now)
            Else
                sAgreementPathTO = sAgreementPathTO & Microsoft.VisualBasic.Month(Now)
            End If

            If Microsoft.VisualBasic.Day(Now) < 10 Then
                sAgreementPathTO = sAgreementPathTO & "0" & Microsoft.VisualBasic.Day(Now)
            Else
                sAgreementPathTO = sAgreementPathTO & Microsoft.VisualBasic.Day(Now)
            End If

            sAgreementPathTO = p_AgreementPath & "\" & sAgreementPathTO

            If (Not System.IO.Directory.Exists(sAgreementPathTO)) Then
                System.IO.Directory.CreateDirectory(sAgreementPathTO)
            End If

            sAgreementPathTO = sAgreementPathTO & "\" & "AgreementTO_" & sPnCID & "_" & sAgreeTime & ".doc"

            System.IO.File.Copy(sTemplate, sAgreementPathTO)

            ObjWord = CreateObject("Word.Application")
            ObjWord.Documents.Open(sAgreementPathTO)

            'ObjWord.Application.Visible = True

            Call ReplaceValues(ObjWord, "<CusName>", sCustName)
            Call ReplaceValues(ObjWord, "<CBBID>", sCBBID)
            Call ReplaceValues(ObjWord, "<AmountBB>", sAmountBB)
            Call ReplaceValues(ObjWord, "<ExpDat>", sExpDate)

            ''StockTable
            Dim r As Integer, c As Integer
            oDoc = ObjWord.ActiveDocument
            oDoc.Range.NoProofing = True

            oTable = oDoc.Tables.Add(oDoc.Bookmarks("STOCKTABLE").Range, iStock, 3)
            oTable.Borders.InsideLineStyle = 0
            oTable.Borders.OutsideLineStyle = 0
            oTable.Range.ParagraphFormat.SpaceAfter = 0
            oTable.Range.ParagraphFormat.LineSpacing = 12
            oTable.Range.Font.Size = 11
            For r = 1 To iStock
                For c = 1 To 3
                    oTable.Cell(r, c).Range.Text = arrStockSold(r - 1, c - 1) 'Data
                Next
            Next
            oTable.Rows(1).Range.Font.Bold = True
            'Table.Rows(1).range.Font.Italic = True
            oTable.Columns(1).Width = ObjWord.InchesToPoints(1.33)
            oTable.Columns(2).Width = ObjWord.InchesToPoints(3.13)
            oTable.Columns(3).Width = ObjWord.InchesToPoints(1.63)


            ObjWord.Application.Visible = True
            ObjWord.ActiveDocument.Save()
            ObjWord = Nothing

        Catch ex As Exception
            ObjWord = Nothing
            System.Windows.Forms.MessageBox.Show("Error in creating Agreement TearOff document." & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub


    Private Sub ReplaceValues(oWord As Object, Header As String, Data As String)
        Dim i As Integer
        Dim strData As String

        Clipboard.Clear()

        With oWord.Selection.Find
            .ClearFormatting()
            .Text = Header
            .Execute(Forward:=True)
        End With

        If Data = "" Then
            Data = " - "
        End If

        strData = ""
        'Data = Replace(Data, vbCrLf, " ")
        For i = 1 To Len(Data)
            'MsgBox Mid(Data, i, 1) & " - " & Asc(Mid(Data, i, 1))
            If Mid(Data, i, 1) <> vbCrLf Then strData = strData + Mid(Data, i, 1)
        Next i

        Clipboard.SetText(CStr(strData))
        oWord.Selection.Paste()
        Clipboard.Clear()

    End Sub
    Private Sub TemplateError()

    End Sub

End Module

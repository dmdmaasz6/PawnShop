Attribute VB_Name = "modTemplates"
Option Explicit
Public arrStockSold(20, 3) As String
Public iStock As Integer


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

If FolderExists(sTemplatePath) <> True Then
    sMsg = "Template Path " & sTemplatePath & " not found:"
    GoTo lblErr
End If

If FolderExists(sTemplatePath & "\" & sTemplateShop) <> True Then
    sMsg = "Template Path " & sTemplatePath & "\" & sTemplateShop & " not found:"
    GoTo lblErr
End If

If FileExists(sTemplatePath & "\" & sTemplateShop & "\" & sTemplate) <> True Then
    sMsg = "Template " & sTemplatePath & "\" & sTemplateShop & "\" & sTemplate & " not found:"
    GoTo lblErr
End If


ValidTemplates = True

Exit Function
lblErr:
MsgBox sMsg & vbCrLf & vbCrLf & "Please correct the error!", vbCritical + vbOKOnly, "Template Error!"

End Function

Public Function Template_Agreement(sCBBID As String, _
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

sTemplate = "agreement.doc"

If ValidTemplates(sTemplate) <> True Then
    TemplateError
    Exit Function
Else
    sTemplate = sTemplatePath & "\" & sTemplateShop & "\" & sTemplate
End If

Set ObjWord = CreateObject("Word.Application")
ObjWord.Documents.Open (sTemplate)

'ObjWord.Application.Visible = True

Call ReplaceValues("<CBBID>", sCBBID)
Call ReplaceValues("<MONTHNR>", sMonthNR)
Call ReplaceValues("<PTTTIME>", sPTTTime)
Call ReplaceValues("<PNCID>", sPnCID)

Call ReplaceValues("<CusName>", sCustName)
Call ReplaceValues("<CusAdr>", sCustAddr)
Call ReplaceValues("<CusID>", sCustID)
Call ReplaceValues("<CusTelW>", sCustTelW)
Call ReplaceValues("<CusTelH>", sCustTelH)
'Call ReplaceValues("<CusAdr>", sCustAddr)
Call ReplaceValues("<AmountP>", sAmountP)
Call ReplaceValues("<AmountBB>", sAmountBB)
Call ReplaceValues("<HandInDat>", sHandInDate)
Call ReplaceValues("<ExpDat>", sExpDate)

'StockTable
   Dim r As Integer, c As Integer
    Set oDoc = ObjWord.ActiveDocument
    Set oTable = oDoc.Tables.Add(oDoc.Bookmarks("STOCKTABLE").range, iStock, 3)
        oTable.Borders.InsideLineStyle = 0
        oTable.Borders.OutsideLineStyle = 0
        oTable.range.ParagraphFormat.SpaceAfter = 0
        oTable.range.ParagraphFormat.LineSpacing = 12
        oTable.range.Font.Size = 9
        For r = 1 To iStock
            For c = 1 To 3
                oTable.Cell(r, c).range.Text = arrStockSold(r - 1, c - 1) 'Data
            Next
        Next
        oTable.Rows(1).range.Font.Bold = True
        'Table.Rows(1).range.Font.Italic = True
        oTable.Columns(1).Width = ObjWord.InchesToPoints(1.33)
        oTable.Columns(2).Width = ObjWord.InchesToPoints(3.13)
        oTable.Columns(3).Width = ObjWord.InchesToPoints(1.63)

ObjWord.Application.Visible = True

ObjWord.ActiveDocument.SaveAs (sTemplatePath & "\NewAgreement.doc")
'ObjWord.ActiveDocument.PrintOut False, _
'                            False, _
'                            , _
'                            sTemplatePath & "\NewAgreement.doc", _
'                            , _
'                            , _
'                            , _
'                            , _
'                            , _
'                            , _
'                            True

'ObjWord.Quit
Set ObjWord = Nothing
    
End Function

Public Function Template_AgreementTearOff(sCBBID As String, _
                                          sCustName As String, _
                                          sAmountBB As String, _
                                          sExpDate As String)
Dim sTemplate As String

sTemplate = "agreement_2.doc"

If ValidTemplates(sTemplate) <> True Then
    TemplateError
    Exit Function
Else
    sTemplate = sTemplatePath & "\" & sTemplateShop & "\" & sTemplate
End If

Set ObjWord = CreateObject("Word.Application")
ObjWord.Documents.Open (sTemplate)

'ObjWord.Application.Visible = True

Call ReplaceValues("<CusName>", sCustName)
Call ReplaceValues("<CBBID>", sCBBID)
Call ReplaceValues("<AmountBB>", sAmountBB)
Call ReplaceValues("<ExpDat>", sExpDate)

'StockTable
   Dim r As Integer, c As Integer
    Set oDoc = ObjWord.ActiveDocument
    Set oTable = oDoc.Tables.Add(oDoc.Bookmarks("STOCKTABLE").range, iStock, 3)
        oTable.Borders.InsideLineStyle = 0
        oTable.Borders.OutsideLineStyle = 0
        oTable.range.ParagraphFormat.SpaceAfter = 0
        oTable.range.ParagraphFormat.LineSpacing = 12
        oTable.range.Font.Size = 11
        For r = 1 To iStock
            For c = 1 To 3
                oTable.Cell(r, c).range.Text = arrStockSold(r - 1, c - 1) 'Data
            Next
        Next
        oTable.Rows(1).range.Font.Bold = True
        'Table.Rows(1).range.Font.Italic = True
        oTable.Columns(1).Width = ObjWord.InchesToPoints(1.33)
        oTable.Columns(2).Width = ObjWord.InchesToPoints(3.13)
        oTable.Columns(3).Width = ObjWord.InchesToPoints(1.63)


ObjWord.Application.Visible = True
ObjWord.ActiveDocument.SaveAs (sTemplatePath & "\NewAgreement_2.doc")
'ObjWord.ActiveDocument.PrintOut False, _
'                            False, _
'                            , _
'                            sTemplatePath & "\NewAgreement.doc", _
'                            , _
'                            , _
'                            , _
'                            , _
'                            , _
'                            , _
'                            True

'ObjWord.Quit
Set ObjWord = Nothing
    
End Function


Private Sub ReplaceValues(Header As String, Data As String)
Dim i As Integer
Dim strData As String

Clipboard.Clear

With ObjWord.Selection.Find
    .ClearFormatting
    .Text = Header
    .Execute Forward:=True
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
        
Clipboard.SetText (CStr(strData))
ObjWord.Selection.Paste
Clipboard.Clear

End Sub
Private Function TemplateError()

End Function

Public Function ClearStockList()
Dim i As Integer

iStock = 0
For i = 0 To 20
    arrStockSold(i, 0) = ""
    arrStockSold(i, 1) = ""
    arrStockSold(i, 2) = ""
    arrStockSold(i, 3) = ""
Next i

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Updater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub updateBookLinks(wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim listNum, firstNum, sheetIndex
    Dim listName, preName, nextName
    Dim i, j, cIndex, sheetsCount
    
    listNum = mUtil.getListSheetNum(wBook)
    firstNum = mUtil.getFirstSheetNum(wBook)
    listName = wBook.Sheets(listNum).Name
'    Debug.Print listNum & "?" & firstNum

    If wBook.Sheets.Count > 159 Then
        Application.ScreenUpdating = False
    End If
    
    With wBook
        sheetsCount = .Sheets.Count
        For i = .Sheets.Count To firstNum Step -1
            Call updateSheetName(.Sheets(i))
            Call updateSheetTab(.Sheets(i))
        Next
        
            'format first
        cIndex = 1
        Call updateSheetLink(.Sheets(firstNum), cIndex, listName, True, False)
        
        If .Sheets.Count >= 3 Then
            For j = firstNum + 1 To .Sheets.Count - 1
                'cIndex
                cIndex = j - firstNum + 1
                Call updateSheetLink(.Sheets(j), cIndex, listName, False, False)
            Next
                'format last Sheet
            cIndex = sheetsCount - firstNum + 1
            Call updateSheetLink(.Sheets(sheetsCount), cIndex, listName, False, True)
        End If
    End With
    
    Set mUtil = Nothing
End Sub
Public Sub updateBookTabs(wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim firstNum, i
    
'    Debug.Print "update tabs"
    firstNum = mUtil.getFirstSheetNum(wBook)
    With wBook
        For i = firstNum To .Sheets.Count
            Call updateSheetTab(.Sheets(i))
        Next
    End With
    
    Set mUtil = Nothing
    
End Sub
Public Sub updateBookHome(wBook As Workbook)
    Dim issueCount, firstNum, linkNum, listNum, sheetsCount
    Dim mUtil As New XlUtil
    Dim indexArr, dirArr
    Dim i, j
    
    On Error Resume Next
    
    firstNum = mUtil.getFirstSheetNum(wBook)
    listNum = mUtil.getListSheetNum(wBook)
    issueCount = wBook.Sheets.Count - firstNum + 1
    sheetsCount = wBook.Sheets.Count
    linkNum = 3
    
    wBook.Sheets(listNum).Activate
    
    If sheetsCount > 159 Then
        Application.ScreenUpdating = False
    End If
    
    indexArr = Array("序号", "问题编号", "问题来源", "问题描述", "所属系统", "提出部门", "提出人", "提出时间", _
    "责任部门", "责任人", "最终措施", "当前进展", "计划完成时间", "实际完成时间", "状态", "问题类别", "提出部门措施", "责任部门措施", "备注")
    dirArr = Array("C3", "C5", "B3", "C7", "C11", "D3", "C9", _
    "D5", "D7", "E9", "E11", "D9", "D11", "E3", "F3", "E5", "E7")
    
    With wBook.Sheets(listNum)
        .AutoFilterMode = False
        
        For i = 0 To UBound(indexArr)
            .Cells(linkNum - 1, i + 1).Value = indexArr(i)
        Next

        For j = linkNum To issueCount + linkNum - 1
            .Cells(j, 1).Value = "=HYPERLINK(""#'("" &ROW(A" & j - 2 & ")& "")'!A1"",ROW(A" & j - 2 & "))"
'            .Cells(j, 1).Value = "=HYPERLINK(""#'("" &ROW(A" & j - 2 & ")& "")'!A1"",INDIRECT(""'(""&ROW(A" & j - 2 & ")&"")'!A3""))"
            For i = 0 To UBound(indexArr)
                .Cells(j, i + 2).Value = "=INDIRECT(""'(""&ROW(A" & j - 2 & ")&"")'!" & dirArr(i) & """)"
            Next
        Next

'        .Range(.Cells(linkNum, 1), .Cells(linkNum + 1, 17)).AutoFill Destination:=.Range(.Cells(linkNum, 1), .Cells(issueCount + linkNum - 1, 17)), Type:=xlFillDefault

'           =HYPERLINK("#'("&ROW(A1)&")'!A1",INDIRECT("'("&ROW(A1)&")'!$a$3"))
'           "=HYPERLINK(""#'(""&ROW(A1)&"")'!A1"",INDIRECT(""'(""&ROW(A1)&"")'!$a$3""))"
       
        With .Range(.Cells(linkNum - 1, 1), .Cells(issueCount + linkNum - 1, 19))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With

        With .Columns("D:D")
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With

        With .Columns("K:M")
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
'        Debug.Print .UsedRange.Rows.Count & "?" & issueCount + linkNum - 1
        If .UsedRange.Rows.Count + linkNum - 1 > issueCount + linkNum - 1 Then
            With .Rows((issueCount + linkNum) & ":" & .UsedRange.Rows.Count + linkNum - 1)
                .ClearContents
                .ClearFormats
            End With
        End If
        
        With .Columns("O:O")
            .FormatConditions.Add Type:=xlTextString, String:="红", TextOperator:=xlContains
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = 255
            .FormatConditions(1).StopIfTrue = False
        End With
        
        With .Columns("O:O")
            .FormatConditions.Add Type:=xlTextString, String:="黄", TextOperator:=xlContains
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = 65535
            .FormatConditions(1).StopIfTrue = False
        End With
        
        With .Columns("O:O")
            .FormatConditions.Add Type:=xlTextString, String:="绿", TextOperator:=xlContains
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = 5287936
            .FormatConditions(1).StopIfTrue = False
        End With
        
        .Columns("H:H").NumberFormatLocal = "yyyy/mm/dd"
        .Columns("M:M").NumberFormatLocal = "yyyy/mm/dd"
        .Columns("N:N").NumberFormatLocal = "yyyy/mm/dd"
        
        .Rows(linkNum - 1).AutoFilter

    End With

    Set mUtil = Nothing

End Sub
Private Sub updateSheetName(wSheet As Worksheet)
    Dim mUtil As New XlUtil
    With wSheet
'        Debug.Print .CodeName & "/" & .Name
'        .Name = mUtil.hashCode(.CodeName & .Name)
        .Name = .CodeName & .Name
    End With
    Set mUtil = Nothing
End Sub
Public Sub updateSheetTab(wSheet As Worksheet)

    With wSheet
        Select Case .Cells(3, 5).Value          'update tabs color
             Case "绿":  .Tab.Color = 5287936
             Case "红":  .Tab.Color = 255
             Case "黄":  .Tab.Color = 65535
             Case Else:  .Tab.ColorIndex = xlNone: .Tab.TintAndShade = 0
        End Select
    End With
End Sub

Public Sub updateSheetLink(wSheet As Worksheet, showIndex As Variant, listSheetName As Variant, isFirst As Boolean, isLast As Boolean)
    Dim preSheetName, nextSheetName
    
    With wSheet
        .Cells.Hyperlinks.Delete
        
        .Name = "(" & showIndex & ")"
        preSheetName = "(" & showIndex - 1 & ")"
        nextSheetName = "(" & showIndex + 1 & ")"
        
         With .Cells(1, 1)       'to homepage
'             .Hyperlinks.Delete
'             .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="'" & listSheetName & "'!A" & (showIndex + 2), _
'                 TextToDisplay:="问题报告", ScreenTip:="单击返回首页"
             .Value = "=HYPERLINK(""#'" & listSheetName & "'!A" & showIndex + 2 & """,""问题报告"")"
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
             .NumberFormatLocal = "G/通用格式"
             .Font.ThemeColor = xlThemeColorLight1
             .Font.Size = 14
             .Font.Bold = True
'                .Font.Underline = xlUnderlineStyleNone
'            With .Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'            End With
'            With .Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'            End With
'            With .Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlThin
'            End With
'            With .Borders(xlEdgeRight)
'                .LineStyle = xlNone
'                .Weight = xlThin
'            End With
         End With
        

        With .Cells(2, 1)         'to pre page
            If isFirst = False Then
'             .Value = "序号"
'                 .Hyperlinks.Delete
'                 .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="'" & preSheetName & "'!A2", _
'                     ScreenTip:="单击跳转上一页"
             .Value = "=HYPERLINK(""#'" & preSheetName & "'!A2"",""序号"")"
            End If
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "G/通用格式"
            .Font.ThemeColor = xlThemeColorLight1
            .Font.Size = 11
            .Font.Bold = False
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End With
        
        With .Cells(2, 3)       'to next page
            If isLast = False Then
'                .Value = "问题编号"
    '                     .Hyperlinks.Delete
    '                     .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="'" & nextSheetName & "'!C2", _
    '                         ScreenTip:="单击跳转下一页"
               .Value = "=HYPERLINK(""#'" & nextSheetName & "'!C2"",""问题编号"")"
            End If
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "G/通用格式"
            .Font.ThemeColor = xlThemeColorLight1
            .Font.Size = 11
            .Font.Bold = False
           With .Borders(xlEdgeLeft)
               .LineStyle = xlContinuous
               .Weight = xlThin
           End With
           With .Borders(xlEdgeTop)
               .LineStyle = xlContinuous
               .Weight = xlThin
           End With
           With .Borders(xlEdgeBottom)
               .LineStyle = xlContinuous
               .Weight = xlThin
           End With
           With .Borders(xlEdgeRight)
               .LineStyle = xlContinuous
               .Weight = xlThin
           End With
        End With
         
        With .Cells(3, 1)         'rewrite sheetIndex
'            .Value = showIndex
'            .Value = "=MID(CELL(""filename"",C1),FIND(""("",CELL(""filename"",C1))+1,FIND("")"",CELL(""filename"",C1))-FIND(""("",CELL(""filename"",C1))-1)"
            .Value = "=HYPERLINK(MID(CELL(""filename"",C1),FIND(""["",CELL(""filename"",C1)),FIND(""]"",CELL(""filename"",C1))-FIND(""["",CELL(""filename"",C1))+1)&""问题汇总清单!A""&(MID(CELL(""filename"",C1),FIND(""("",CELL(""filename"",C1))+1,FIND("")"",CELL(""filename"",C1))-FIND(""("",CELL(""filename"",C1))-1)+2),MID(CELL(""filename"",C1),FIND(""("",CELL(""filename"",C1))+1,FIND("")"",CELL(""filename"",C1))-FIND(""("",CELL(""filename"",C1))-1))"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "G/通用格式"
            .Font.ThemeColor = xlThemeColorLight1
            .Font.Size = 11
            .Font.Bold = False
               With .Borders(xlEdgeLeft)
                   .LineStyle = xlContinuous
                   .Weight = xlMedium
               End With
               With .Borders(xlEdgeTop)
                   .LineStyle = xlContinuous
                   .Weight = xlThin
               End With
               With .Borders(xlEdgeBottom)
                   .LineStyle = xlContinuous
                   .Weight = xlThin
               End With
               With .Borders(xlEdgeRight)
                   .LineStyle = xlContinuous
                   .Weight = xlThin
               End With
        End With
        
        With .Range(.Cells(1, 1), .Cells(3, 5))
'            With .Borders(xlInsideVertical)
'                .LineStyle = xlContinuous
'                .Weight = xlThin
'            End With
'            With .Borders(xlInsideHorizontal)
'                .LineStyle = xlContinuous
'                .Weight = xlThin
'            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
'            With .Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'            End With
'            With .Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'            End With
        End With
         
    End With

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Updater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub updateBookTabs(wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim firstNum, i
    
'    Debug.Print "update tabs"
    firstNum = mUtil.getFirstSheetNum(wBook)
    With wBook
        For i = firstNum To .Sheets.Count
            Call updateTabColor(.Sheets(i))
        Next
    End With
    
    Set mUtil = Nothing
    
End Sub
Public Sub updateBookLinks(wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim listNum, firstNum, sheetIndex
    Dim listName, preName, nextName
    Dim i, j, cIndex, sheetsCount
    
    listNum = mUtil.getListSheetNum(wBook)
    firstNum = mUtil.getFirstSheetNum(wBook)
    listName = wBook.Sheets(listNum).Name
'    Debug.Print listNum & "?" & firstNum
    
    With wBook
        sheetsCount = .Sheets.Count
        For i = .Sheets.Count To firstNum Step -1
            Call updateSheetName(.Sheets(i))
        Next
        
            'format first
        cIndex = 1
        Call updateSheetLink(.Sheets(sheetsCount), cIndex, listName, True, False)
    
        For j = firstNum + 1 To .Sheets.Count - 1
            'cIndex
            cIndex = j - firstNum + 1
            Call updateSheetLink(.Sheets(j), cIndex, listName, False, False)
        Next
        'format last Sheet
        cIndex = sheetsCount - firstNum + 1
        Call updateSheetLink(.Sheets(sheetsCount), cIndex, listName, False, True)
    End With
    
    Set mUtil = Nothing
End Sub
Private Sub updateSheetName(wSheet As Worksheet)
    Dim mUtil As New XlUtil
    With wSheet
        Debug.Print .CodeName & "/" & .Name
        .Name = mUtil.hashCode(.CodeName & .Name)
    End With
    Set mUtil = Nothing
End Sub
Public Sub updateTabColor(wSheet As Worksheet)

    With wSheet
        Select Case .Cells(3, 5).Value          'update tabs color
             Case "绿":  .Tab.Color = 5287936
             Case "红":  .Tab.Color = 255
             Case "黄":  .Tab.Color = 65535
             Case Else:  .Tab.ColorIndex = xlNone: .Tab.TintAndShade = 0
        End Select
    End With
End Sub

Private Sub updateSheetLink(wSheet As Worksheet, showIndex As Variant, listSheetName As Variant, isFirst As Boolean, isLast As Boolean)
    Dim preSheetName, nextSheetName
    
    Call updateTabColor(wSheet)
    
    With wSheet
        .Cells.Hyperlinks.Delete
        
        .Name = "(" & showIndex & ")"
        preSheetName = "(" & showIndex - 1 & ")"
        nextSheetName = "(" & showIndex + 1 & ")"
        
         With .Cells(1, 1)       'to homepage
'             .Hyperlinks.Delete
             .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="'" & listSheetName & "'!A" & (showIndex + 2), _
                 TextToDisplay:="问题报告", ScreenTip:="单击返回首页"
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
             .NumberFormatLocal = "G/通用格式"
             .Font.ThemeColor = xlThemeColorLight1
             .Font.Size = 14
             .Font.Bold = True
'                .Font.Underline = xlUnderlineStyleNone
         End With
        
        If isFirst = False Then
             With .Cells(2, 1)         'to pre page
                 .Value = "序号"
    '             .Hyperlinks.Delete
                 .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="'" & preSheetName & "'!A2", _
                     ScreenTip:="单击跳转上一页"
                 .HorizontalAlignment = xlCenter
                 .VerticalAlignment = xlCenter
                 .NumberFormatLocal = "G/通用格式"
                 .Font.ThemeColor = xlThemeColorLight1
                 .Font.Size = 11
                 .Font.Bold = False
             End With
        End If
        
         If isLast = False Then
                 With .Cells(2, 3)       'to next page
                     .Value = "问题序号"
'                     .Hyperlinks.Delete
                     .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="'" & nextSheetName & "'!C2", _
                         ScreenTip:="单击跳转下一页"
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlCenter
                     .NumberFormatLocal = "G/通用格式"
                    .Font.ThemeColor = xlThemeColorLight1
                     .Font.Size = 11
                     .Font.Bold = False
                 End With
         End If
    End With

End Sub

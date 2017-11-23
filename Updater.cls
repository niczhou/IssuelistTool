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
Public Sub updateTabs(wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim firstNum, i
    
    Debug.Print "update tabs"
    firstNum = mUtil.getFirstSheetNum(wBook)
    With wBook
        For i = firstNum To .Sheets.Count
            Call updateTabColor(.Sheets(i))
        Next
    End With
    
    Set mUtil = Nothing
    
End Sub
Public Sub updateLinks(wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim firstNum, i
    
    Debug.Print "update tabs"
    firstNum = mUtil.getFirstSheetNum(wBook)
    With wBook
        For i = firstNum To .Sheets.Count
            Call updateSheetLink(.Sheets(i), wBook)
        Next
    End With
    
    Set mUtil = Nothing
End Sub
Public Sub updateList()

End Sub

Public Sub updateTabColor(wSheet As Worksheet)

    With wSheet
        Select Case .Cells(3, 5).Value          '根据问题状态更新sheet标签颜色
             Case "绿":  .Tab.Color = 5287936
             Case "红":  .Tab.Color = 255
             Case "黄":  .Tab.Color = 65535
             Case Else:  .Tab.ColorIndex = xlNone: .Tab.TintAndShade = 0
        End Select
    End With
End Sub

Public Sub updateSheetLink(wSheet As Worksheet, wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim listNum, firstNum, sheetIndex
    Dim listName, preName, nextName
    
    firstNum = mUtil.getFirstSheetNum(wBook)
    listNum = mUtil.getListSheetNum(wBook)
    listName = wBook.Sheets(listNum).Name
    sheetIndex = mUtil.getSheetIndex(wSheet, wBook)
    preName = wBook.Sheets(sheetIndex - 1).Name
    nextName = wBook.Sheets(sheetIndex + 1).Name
    
    
    With wSheet
        .Name = "(" & (sheetIndex - firstNum + 1) & ")"
            
         Select Case .Cells(3, 5).Value          '根据问题状态更新sheet标签颜色
             Case "绿":  .Tab.Color = 5287936
             Case "红":  .Tab.Color = 255
             Case "黄":  .Tab.Color = 65535
             Case Else:  .Tab.ColorIndex = xlNone: .Tab.TintAndShade = 0
        End Select
        
         With .Cells(1, 1)       '返回首页
             .Hyperlinks.Delete
             .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="'" & listName & "'!A" & (sheetIndex - firstNum + listNum + 1), _
                 TextToDisplay:="问题报告", ScreenTip:="单击返回首页"
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
             .NumberFormatLocal = "G/通用格式"
             .Font.ThemeColor = xlThemeColorLight1
             .Font.Size = 14
             .Font.Bold = True
'                .Font.Underline = xlUnderlineStyleNone
         End With

         With .Cells(3, 1)         '上一页
             .Value = sheetIndex - firstNum + 1
             .Hyperlinks.Delete
             .Hyperlinks.Add Anchor:=.Cells(3, 1), Address:="", SubAddress:="'" & preName & "'!A3", _
                 ScreenTip:="单击跳转上一页"
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
             .NumberFormatLocal = "G/通用格式"
'                .Font.ThemeColor = xlThemeColorLight1
             .Font.Size = 11
             .Font.Bold = False
         End With
        
         If sheetIndex < wBook.Sheets.Count Then
                 With .Cells(3, 3)       '下一页
                     .Hyperlinks.Delete
                     .Hyperlinks.Add Anchor:=Sheets(i).Cells(3, 3), Address:="", SubAddress:="'" & nextName & "'!C3", _
                         ScreenTip:="单击跳转下一页"
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlCenter
                     .NumberFormatLocal = "G/通用格式"
'                        .Font.ThemeColor = xlThemeColorLight1
                     .Font.Size = 11
                     .Font.Bold = False
                 End With
         End If
    End With
End Sub

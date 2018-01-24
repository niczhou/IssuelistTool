VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim inputArr
Public Function inputBook(wBook As Workbook, startSheetNum As Variant, endSheetNum As Variant, strItem As Variant, strInput As Variant)
    Dim i, j
    Dim mCreater As Creater
    
    With wBook
        For i = startSheetNum To endSheetNum
            If strItem = "问题编号前缀" Then
                j = j + 1
                .Sheets(i).Cells(3, 3) = strInput & "-" & j
            Else
                Call InputSheet(.Sheets(i), strItem, strInput)
            End If
        Next
    End With
    
    Application.DisplayAlerts = False
'    If mCreater <> Nothing Then
        Set mCreater = Nothing
'    End If

End Function
Public Sub inputBookByFilter(wBook As Workbook, strItem As Variant, strInput As Variant)
    Dim mUtil As New XlUtil, mCreater As Creater
    Dim listNum, cRow, iSht, usedCount
    Dim filtArr()
    Dim i, j, i_filt
    
    listNum = mUtil.getListSheetNum(wBook)
    filtArr = mUtil.getFiltArr(wBook.Sheets(listNum))
    
    With wBook
        For i = 0 To UBound(filtArr)
'            Debug.Print .Sheets(listNum).Cells(filtArr(i), 1).Value
            i_filt = .Sheets(listNum).Cells(filtArr(i), 1).Value
            If i_filt <> "" And IsNumeric(i_filt) = True And i_filt < .Sheets.Count Then
                If strItem = "问题编号前缀" Then
                    j = j + 1
                    .Sheets("(" & i_filt & ")").Cells(3, 3) = strInput & "-" & j
                Else
                    Call InputSheet(.Sheets("(" & i_filt & ")"), strItem, strInput)
                End If
            End If
        Next
    End With
    
    Application.DisplayAlerts = False
'    If mCreater <> Nothing Then
        Set mCreater = Nothing
'    End If
    
End Sub
Public Sub initInputFiter(mComboBox As ComboBox, strTip As Variant)
    Dim i
    
    inputArr = Array("问题编号前缀", "问题来源", "提出人", "提出部门", "所属系统", "提出时间", "计划完成时间", "实际完成时间", _
     "责任人", "责任部门", "问题类别", "项目", "格式化工作表")
     
    mComboBox.Style = fmStyleDropDownList
    mComboBox.Clear
    mComboBox.AddItem strTip
    mComboBox.ListIndex = 0
    For i = 0 To UBound(inputArr)
        mComboBox.AddItem inputArr(i)
    Next
End Sub
Private Function InputSheet(wSheet As Worksheet, strItem As Variant, strInput As Variant)
    Dim cl As Range
    
    With wSheet
        If strItem = "格式化工作表" Then
            Application.DisplayAlerts = False
            If .UsedRange.Columns.Count = 6 Then
                .Columns("A:A").ColumnWidth = 8
                .Columns("B:B").ColumnWidth = 65
                .Columns("C:C").ColumnWidth = 11.5
                .Columns("D:D").ColumnWidth = 11.5
                .Columns("E:E").ColumnWidth = 17
                .Columns("F:F").ColumnWidth = 17
            ElseIf .UsedRange.Columns.Count = 5 Then
                .Columns("A:A").ColumnWidth = 8
                .Columns("B:B").ColumnWidth = 72
                .Columns("C:C").ColumnWidth = 15
                .Columns("D:D").ColumnWidth = 15
                .Columns("E:E").ColumnWidth = 20
            End If
            
            .Rows("1:1").RowHeight = 20
            .Rows("2:2").RowHeight = 20
            .Rows("3:3").RowHeight = 48
            .Rows("4:4").RowHeight = 20
            .Rows("5:5").RowHeight = 72.5
            .Rows("6:6").RowHeight = 20
            .Rows("7:7").RowHeight = 72.5
            .Rows("8:8").RowHeight = 20
            .Rows("9:9").RowHeight = 72.5
            .Rows("10:10").RowHeight = 20
            .Rows("11:11").RowHeight = 72.5
        ElseIf strInput <> "" Then
            If strItem = "问题编号前缀" Then
'                .Cells(3, 3).Value = strInput & "-"
            ElseIf strItem = "项目" Then
                .Cells(1, 6).Value = strInput
            Else
                For Each cl In .UsedRange
                    If cl.Value Like strItem Then
    '                    Debug.Print cl.Row & "/" & cl.Column
                        cl.Offset(1, 0).Value = strInput
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
End Function

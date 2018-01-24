VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Deleter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub deleteBook(mBook As Workbook, startNum, endNum)
    Dim i
    For i = endNum To startNum Step -1
        mBook.Sheets("(" & i & ")").Delete
    Next

End Sub
Public Sub deleteBookByFilter(wBook As Workbook)
    Dim mUtil As New XlUtil
    Dim listNum
    Dim filtArr()
    Dim cl As Range
    Dim i, i_filt, i_index
    
'    On Error GoTo ErrorHandler
    
    listNum = mUtil.getListSheetNum(wBook)
    filtArr = mUtil.getFiltArr(wBook.Sheets(listNum))
'    Debug.Print UBound(filtArr)
    
    With wBook
        For i = UBound(filtArr) To 0 Step -1
'            Debug.Print .Sheets(listNum).Cells(filtArr(i), 1).Value
            i_filt = .Sheets(listNum).Cells(filtArr(i), 1).Value
            If i_filt <> "" And IsNumeric(i_filt) = True And i_filt < .Sheets.Count Then
                .Sheets("(" & i_filt & ")").Delete
            End If
        Next
    End With
            
    Set mUtil = Nothing
    
'ErrorHandler:
'    Exit Sub
'    Set mUtil = Nothing
'    MsgBox "要删除的Sheet不存在，请更新清单后重试!"
End Sub
Public Sub deleteSheet(sheetItem, wBook As Workbook, Optional isNumericMode As Boolean = False)
    Dim mUtil As New XlUtil
    
    If isNumericMode = False Then
        If mUtil.sheetIndex(sheetItem, wBook) > 0 Then
            wBook.Sheets("""" & sheetItem & """").Delete
        Else
            MsgBox "错误404：工作表" & sheetItem & " 不存在！"
        End If
    Else
        If IsNumeric(sheetItem) = True And sheetItem < wBook.Sheets.Count Then
            wBook.Sheets(sheetItem).Delete
        Else
            MsgBox "错误504：工作表" & sheetItem & " 数值大于Excel工作表数量！"
        End If
    End If
End Sub

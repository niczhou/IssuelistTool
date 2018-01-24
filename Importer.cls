VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Importer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub importBook(fromBook As Workbook, toBook As Workbook, startNum As Variant, endNum As Variant, afterNum As Variant)
    Dim i

    For i = startNum To endNum
'    fromBook startNum endNum=> name number --to avoid wrong input
'    toBook afterNum=>count number --to work without alert
        fromBook.Sheets("(" & i & ")").Copy after:=toBook.Sheets(afterNum)
        afterNum = afterNum + 1
    Next

End Sub

Public Sub importBookByFilter(fromBook As Workbook, toBook As Workbook, afterNum As Variant)
    Dim mUtil As New XlUtil
    Dim oList, usedCount
    Dim i, i_filt, i_index
    Dim filtArr()
    Dim cl As Range
    
    oList = mUtil.getListSheetNum(fromBook)
    filtArr = mUtil.getFiltArr(fromBook.Sheets(oList))
    
    For i = 0 To UBound(filtArr)
        With fromBook
'            Debug.Print .Sheets(listNum).Cells(filtArr(i), 1).Value
            i_filt = .Sheets(oList).Cells(filtArr(i), 1).Value
'    fromBook startNum endNum=> name number --to avoid wrong input
'    toBook afterNum=>count number --to work without alert
            If i_filt <> "" And IsNumeric(i_filt) = True And i_filt < .Sheets.Count Then
                .Sheets("(" & i_filt & ")").Copy after:=toBook.Sheets(afterNum)
                afterNum = afterNum + 1
            End If
        End With
    Next
    Set mUtil = Nothing
    
End Sub

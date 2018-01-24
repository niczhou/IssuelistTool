VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim isExcelClosed As Boolean

Public Function hashCode(ByVal str As Variant) As Variant
    Dim i As Integer
    Dim hashTemp
    
    hashTemp = 1

    For i = Len(str) - 1 To 0 Step -1
        hashTemp = Left(Left(31 * hashTemp, 8) + Asc(Mid(str, i + 1, 1)), 8)
    Next
    hashCode = hashTemp
End Function
Public Function sheetCode(wSheet As Worksheet) As Variant
    sheetCode = Replace(wSheet.CodeName, "Sheet", "") & wSheet.Name
End Function
Public Function indexOfArray(element As Variant, arr As Variant) As Integer
    Dim i
    Dim isExist As Boolean
    
    If IsArray(arr) = False Then
        indexOfArray = -1
        Exit Function
    ElseIf ArrayRange(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            If arr(i) <> element Then
                indexOfArray = -1
            Else
                indexOfArray = i
                Exit For
            End If
        Next
    ElseIf ArrayRange(arr) = 2 Then
        For i = LBound(arr, 2) To UBound(arr, 2)
            If arr(0, i) <> element Then
                indexOfArray = -1
            Else
                indexOfArray = i
                Exit For
            End If
        Next
    End If
End Function
Public Function initUI(ExcelApp As Excel.Application, mComboBox As ComboBox, strTip As String)
    Dim i
    On Error GoTo ErrorHandler
    
    mComboBox.Style = fmStyleDropDownList
    mComboBox.Clear
    mComboBox.AddItem strTip
    mComboBox.AddItem "打开新Excel文件"
    mComboBox.ListIndex = 0
    
ErrorHandler:
    Exit Function
End Function
Public Function initComboBox(ExcelApp As Excel.Application, mComboBox As ComboBox, strTip As String)
    Dim i
    On Error GoTo ErrorHandler
    
    mComboBox.Style = fmStyleDropDownList
    mComboBox.Clear
    mComboBox.AddItem strTip
    mComboBox.AddItem "打开新Excel文件"
    mComboBox.ListIndex = 0
    
    If isExcelClosed = False Then
        If ExcelApp.Workbooks.Count > 0 Then
            For i = 1 To ExcelApp.Workbooks.Count
                If ExcelApp.Workbooks(i).Name <> ThisWorkbook.Name Then
                    mComboBox.AddItem ExcelApp.Workbooks(i).Name
                End If
            Next
        End If
    End If
ErrorHandler:
    Exit Function

End Function

Public Function getBook(xlApp As Excel.Application, bookName As String) As Workbook
    Dim i
    
    On Error GoTo ErrorHandler
    If bookName Like "选择*" Then
        MsgBox "请先选择Excel文件！"
    ElseIf bookName Like "打开新Excel文件" Then
        Set getBook = xlApp.Workbooks.Open(Filename:=xlApp.GetOpenFilename("Excel Files (*.*), Excel Files(*.*)", , "打开新Excel文件"))
    ElseIf isExcelClosed = False Then
        For i = 1 To xlApp.Workbooks.Count
            If xlApp.Workbooks(i).Name Like bookName Then
                Set getBook = xlApp.Workbooks(i)
                Exit Function
            End If
        Next
    End If
ErrorHandler:
    Exit Function
    MsgBox "打开文件错误，请打开正确的Excel文件"
End Function

Public Function getXlApp() As Excel.Application

    On Error GoTo ErrorHandler
    
    Set getXlApp = GetObject(, "Excel.Application")
'    Debug.Print getXlApp.Workbooks.Count
    
ErrorHandler:
    Exit Function
    isExcelClosed = True
    Set getXlApp = CreateObject("Excel.Application")
'    Debug.Print getXlApp.Workbooks.Count
    
End Function

Public Function getListSheetNum(wBook As Workbook) As Integer
    Dim i
    Dim nl
    
    getListSheetNum = 1
    With wBook
        For i = 1 To .Sheets.Count
            If .Sheets(i).Name Like "*清单*" Or .Sheets(i).Name Like "*list*" Then
                getListSheetNum = i
                Exit Function
            End If
        Next
    End With
    
    MsgBox "不能识别问题清单汇总页，请重命名为 '问题汇总清单'"
End Function
Public Function getFirstSheetNum(wBook As Workbook) As Integer
    Dim i
    Dim nf
    
    getFirstSheetNum = 2
    With wBook
        For i = 1 To .Sheets.Count
            If .Sheets(i).Name Like "(#)" Or .Sheets(i).Name Like "*#*" Or .Sheets(i).Name Like "(1)" Then
                getFirstSheetNum = i
                Exit Function
            End If
        Next
    End With
    MsgBox "不能识别问题清单第一个问题，请重命名第一个问题为 '(1)'"
    
End Function
Public Function getSheetIndex(wSheet As Worksheet, wBook As Workbook) As Integer
    Dim i
    Dim ws As Worksheet
    With wBook
        For i = 1 To wBook.Sheets.Count
            If .Sheets(i).Name = wSheet.Name Then
                getSheetIndex = i
                Exit Function
            Else
                getSheetIndex = -1
            End If
        Next
    End With
End Function
Public Function sheetIndex(sheetName As Variant, wBook As Workbook) As Integer
    Dim i
    Dim ws As Worksheet
    With wBook
        For i = 1 To wBook.Sheets.Count
            If .Sheets(i).Name = sheetName Then
                sheetIndex = i
                Exit Function
            Else
                sheetIndex = -1
            End If
        Next
    End With
End Function
Public Function getFiltArr(wSheet As Worksheet) As Variant()
    Dim usedCount, i_hidden, i_row
    Dim filtArr()
    
    With wSheet
        usedCount = .UsedRange.Rows.Count
        ReDim filtArr(0 To usedCount - 1)
        For i_row = 1 To usedCount
'            Debug.Print i_row & ":" & .Rows(i_row).Hidden
            If .Rows(i_row).Hidden = False Then
                filtArr(i_hidden) = i_row
                i_hidden = i_hidden + 1
            End If
        Next
    End With
    
    ReDim Preserve filtArr(0 To i_hidden - 1)
    getFiltArr = filtArr
'    Debug.Print UBound(getFilterArr)
End Function

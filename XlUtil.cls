VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
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
    Debug.Print str
    For i = 0 To Len(str) - 1
        hashTemp = Left(Left(31 * hashTemp, 6) + Asc(Mid(str, i + 1, 1)), 6)
    Next
    Debug.Print hashTemp
    hashCode = hashTemp
End Function

Public Function initComboBox(ExcelApp As Excel.Application, mCombobox As ComboBox, strTip As String)
    Dim i
    On Error GoTo ErrorHandler
    
'    mCombobox.Style = fmStyleDropDownList
    mCombobox.Clear
    mCombobox.AddItem strTip
    mCombobox.AddItem "打开新Excel文件"
    mCombobox.ListIndex = 0
    
    If isExcelClosed = False Then
        If ExcelApp.Workbooks.Count > 0 Then
            For i = 1 To ExcelApp.Workbooks.Count
                mCombobox.AddItem ExcelApp.Workbooks(i).Name
            Next
        End If
    End If
ErrorHandler:
    Exit Function

End Function

Public Function getBook(xlApp As Excel.Application, bookName As String) As Workbook
    Dim i
    
    On Error GoTo ErrorHandler
    If (bookName = "打开新Excel文件") Then
        Set getBook = xlApp.Workbooks.Open(FileName:=xlApp.GetOpenFilename("Excel Files (*.*), Excel Files(*.*)", , "打开新Excel文件"))
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
    
    nl = 1
    With wBook
        For i = 1 To .Sheets.Count
            If .Sheets(i).Name Like "*清单*" Or .Sheets(i).Name Like "*list*" Then
                nl = i
                Exit For
            End If
        Next
    End With
    getListSheetNum = nl
End Function
Public Function getFirstSheetNum(wBook As Workbook) As Integer
    Dim i
    Dim nf
    
    nf = 2
    With wBook
        For i = 1 To .Sheets.Count
            If IsNumeric(.Sheets(i).Name) Then
                nf = i
                Exit For
            End If
        Next
    End With
    getFirstSheetNum = nf
End Function
Public Function getSheetIndex(wSheet As Worksheet, wBook As Workbook) As Integer
    Dim i
    Dim ws As Worksheet
    For i = 1 To wBook.Sheets.Count
        If ws.Name = wSheet.Name Then
            getSheetIndex = i
            Exit Function
        Else
            getSheetIndex = -1
        End If
    Next
End Function
Public Function getFilterArr(wSheet As Worksheet, strColumn As Variant) As Integer()
    Dim filterArr
    Dim rg As Range
    With wSheet
        For Each rg In .Columns(strColumn)
            If rg.Show = True Then
                filterArr.Add (rg.Row)
                Debug.Print rg.Row
            End If
        Next
    End With
    getFilterArr = filterArr
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim inputArr
Public Function inputBook(wBook As Workbook)

End Function

Public Sub initInputFiter(mCombo As ComboBox, strTip As Variant)
    Dim i
    
    inputArr = Array("问题编号前缀", "问题来源", "提出时间", "提出人", "提出部门", "所属系统", _
     "责任人", "责任部门", "项目")
     
    mCombo.AddItem strTip
    mCombo.ListIndex = 0
    For i = 0 To UBound(inputArr)
        mCombo.AddItem inputArr(i)
    Next
End Sub
Private Function getInputCell(strInput)

End Function

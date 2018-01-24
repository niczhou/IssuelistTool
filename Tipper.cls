VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function makeTip(strTask As Variant)
    Dim strTip
    
    Select Case strTask
        Case ""
'            strTip = "欢迎使用问题清单工具！" & Chr(13) & Chr(13) & "请点击左侧菜单选择功能"
            UserForm2.Show 1
            Exit Function
        Case "UPDATE"
           strTip = "1.选择Excel文件" & Chr(13) & "2.选择更新的内容"
        Case "CREATE"
            strTip = "1.选择工作Excel文件" & Chr(13) & "2.选择新问题的位置" & Chr(13) & "3.选择新问题的数量"
        Case "IMPORT"
         strTip = "请先更新问题清单，否则不能识别工作表！" & Chr(13) & Chr(13) & "1.选择导入的Excel文件" & Chr(13) & "2.选择导入的问题范围" & Chr(13) & "3.选择合并的Excel文件" & Chr(13) & "4.选择新问题的位置"
        Case "DELETE"
           strTip = "请先更新问题清单，否则不能识别工作表！" & Chr(13) & Chr(13) & "1.选择Excel文件" & Chr(13) & "2.选择删除范围"
        Case "INPUT"
            strTip = "请先更新问题清单，否则不能识别工作表！" & Chr(13) & Chr(13) & "1.选择Excel文件" & Chr(13) & "2.选择输入项目" & Chr(13) & "3.填写输入内容"
    End Select
    
    MsgBox strTip
End Function

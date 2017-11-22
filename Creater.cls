VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Creater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub formatNewSheet(xlSheet As Worksheet)

    With xlSheet
    
        With .Range(.Cells(1, 1), .Cells(11, 6))
    
            With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        
        .Range(.Cells(1, 1), .Cells(1, 4)).Merge (True)
        .Range(.Cells(4, 1), .Cells(4, 2)).Merge
        .Range(.Cells(4, 5), .Cells(4, 6)).Merge
        .Range(.Cells(5, 1), .Cells(11, 2)).Merge
        .Range(.Cells(5, 5), .Cells(5, 6)).Merge
        .Range(.Cells(6, 5), .Cells(6, 6)).Merge
        .Range(.Cells(7, 5), .Cells(7, 6)).Merge
        .Range(.Cells(8, 5), .Cells(8, 6)).Merge
        .Range(.Cells(9, 5), .Cells(9, 6)).Merge
        .Range(.Cells(10, 5), .Cells(10, 6)).Merge
        .Range(.Cells(11, 5), .Cells(11, 6)).Merge
        
        .Cells(1, 1).Value = "问题报告"
    '        Cells(1, 1).InsertIndent 1
        .Cells(1, 5).Value = "项目："
        .Cells(2, 1).Value = "序号"
        .Cells(2, 1).Value = "序号"
        .Cells(2, 2).Value = "问题描述"
        .Cells(2, 3).Value = "问题编号"
        .Cells(2, 4).Value = "提出人"
        .Cells(2, 5).Value = "问题状态"
        .Cells(2, 6).Value = "问题级别"
        .Cells(4, 1).Value = "详细描述"
        .Cells(4, 3).Value = "问题来源"
        .Cells(4, 4).Value = "责任部门"
        .Cells(4, 5).Value = "提出部门措施"
        .Cells(6, 3).Value = "所属系统"
        .Cells(6, 4).Value = "责任人"
        .Cells(6, 5).Value = "责任部门措施"
        .Cells(8, 3).Value = "提出时间"
        .Cells(8, 4).Value = "计划完成时间"
        .Cells(8, 5).Value = "最终措施"
        .Cells(10, 3).Value = "提出部门"
        .Cells(10, 4).Value = "实际完成时间"
        .Cells(10, 5).Value = "当前进展"
        
        .Columns("A:A").ColumnWidth = 8
        .Columns("B:B").ColumnWidth = 56
        .Columns("C:C").ColumnWidth = 12
        .Columns("D:D").ColumnWidth = 12
        .Columns("E:E").ColumnWidth = 20
        .Columns("F:F").ColumnWidth = 20
        
        .Rows("1:1").RowHeight = 20
        .Rows("2:2").RowHeight = 20
        .Rows("3:3").RowHeight = 48
        .Rows("4:4").RowHeight = 20
        .Rows("5:5").RowHeight = 72
        .Rows("6:6").RowHeight = 20
        .Rows("7:7").RowHeight = 72
        .Rows("8:8").RowHeight = 20
        .Rows("9:9").RowHeight = 72
        .Rows("10:10").RowHeight = 20
        .Rows("11:11").RowHeight = 72
        
        With .Range(.Cells(1, 1), .Cells(11, 6))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With .Cells(1, 1)
            .Font.Bold = True
            .Font.Size = 14
        End With
        
        With .Range(.Cells(1, 1), .Cells(1, 2))
            .Borders(xlInsideVertical).LineStyle = xlNone
        End With
        
        With .Range(.Cells(1, 5), .Cells(1, 6))
            .Font.Bold = False
            .Font.Size = 11
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
        End With
        
        With .Range(.Cells(12, 1), .Cells(12, 6))
            .Merge
            .HorizontalAlignment = xlLeft
            .Value = "编制：                                        审核：                                        批准："
        End With
            
        With .Cells(1, 1)
            .Hyperlinks.Delete
'            .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:="", SubAddress:="问题汇总清单!A" & (Sheets.Count - 1)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.ThemeColor = xlThemeColorLight1
            .Font.Size = 14
            .Font.Bold = True
        End With
        
            With .Cells(3, 2)
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        With .Cells(5, 5)
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        With .Cells(7, 5)
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        With .Cells(9, 5)
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        With .Cells(11, 5)
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
         With .Cells(3, 5).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="红,黄,绿"
            .IgnoreBlank = True
            .InCellDropdown = True
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        With .Cells(3, 5)
            .FormatConditions.Add Type:=xlTextString, String:="红", TextOperator:=xlContains
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = 255
            .FormatConditions(1).StopIfTrue = False
        End With
        
        With .Cells(3, 5)
            .FormatConditions.Add Type:=xlTextString, String:="黄", TextOperator:=xlContains
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = 65535
            .FormatConditions(1).StopIfTrue = False
        End With
        
        With .Cells(3, 5)
            .FormatConditions.Add Type:=xlTextString, String:="绿", TextOperator:=xlContains
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).Interior.Color = 5287936
            .FormatConditions(1).StopIfTrue = False
        End With
        
        With .Cells(3, 6).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="A,B,C"
            .IgnoreBlank = True
            .InCellDropdown = True
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
    
    
        .Cells(3, 2).Select
    End With
End Sub


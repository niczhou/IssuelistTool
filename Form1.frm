VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "问题清单工具"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   8610
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "批量删除"
      Height          =   3555
      Index           =   3
      Left            =   120
      TabIndex        =   41
      Top             =   1080
      Width           =   7080
      Begin VB.Frame Frame14 
         Caption         =   "选择文件"
         Height          =   735
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   6855
         Begin VB.ComboBox Combo6 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "删除的范围"
         Height          =   735
         Left            =   120
         TabIndex        =   42
         Top             =   2040
         Width           =   6855
         Begin VB.OptionButton Option12 
            Caption         =   "删除筛选"
            Height          =   375
            Left            =   2040
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option31 
            Caption         =   "删除所有；"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option31 
            Caption         =   "从"
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   45
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   4320
            TabIndex        =   44
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   5280
            TabIndex        =   43
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "页到"
            Height          =   255
            Left            =   4920
            TabIndex        =   48
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "页；"
            Height          =   255
            Left            =   6000
            TabIndex        =   47
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "导入清单"
      Height          =   3555
      Index           =   2
      Left            =   1320
      TabIndex        =   24
      Top             =   5760
      Width           =   7080
      Begin VB.Frame Frame12 
         Caption         =   "导入的范围"
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   6855
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   3600
            TabIndex        =   39
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2640
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option15 
            Caption         =   "从"
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   36
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option15 
            Caption         =   "导入所有；"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "页；"
            Height          =   255
            Left            =   4320
            TabIndex        =   40
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "页到"
            Height          =   255
            Left            =   3240
            TabIndex        =   38
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "导入到"
         Height          =   735
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   6855
         Begin VB.ComboBox Combo4 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "从此文件导入"
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   6855
         Begin VB.ComboBox Combo3 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "导入后的位置"
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   6855
         Begin VB.OptionButton Option11 
            Caption         =   "在清单最后；"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option11 
            Caption         =   "在"
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   27
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2640
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "页之后；"
            Height          =   255
            Left            =   3240
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   7095
      Begin VB.CommandButton Command2 
         Caption         =   "提示"
         Height          =   495
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Height          =   495
         Left            =   1080
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "更新清单"
      Height          =   2000
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   6240
      Width           =   7080
      Begin VB.Frame Frame6 
         Caption         =   "Frame4"
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   6855
         Begin VB.OptionButton Option10 
            Caption         =   "更新汇总清单"
            Height          =   375
            Left            =   3240
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "仅更新标签"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "更新清单链接"
            Height          =   375
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6855
         Begin VB.ComboBox Combo2 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "创建问题"
      Height          =   2000
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   7080
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   6855
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2640
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "在"
            Height          =   375
            Left            =   2160
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option1 
            Caption         =   "在最后"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "之后"
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6855
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   4455
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择工具"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.OptionButton Option7 
         Caption         =   "更新清单"
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "导入清单"
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "创建问题"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option9 
         Caption         =   "批量输入"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "批量删除"
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim btnID As Variant
Dim xlApp As Excel.Application
Dim frameList(2) As msforms.Control

Private Sub Command2_Click()
'   msg hint
End Sub

Private Sub Form_Initialize()
'    frameList = Array(Frame2, Frame5)
End Sub
Private Sub Form_Load()
'    Debug.Print Frame2.Name
    Option6 = True
    Option15(0) = True
    Me.Height = 6266
    Me.Width = 7515
End Sub
Private Sub Command1_Click()
    handleTask (btnID)
End Sub
Private Sub Option1_Click()
    Option2 = False
End Sub
Private Sub Option2_Click()
    Option1 = False
End Sub
Private Sub Option5_Click()
    initTask ("CREATE")
End Sub

Private Sub Option6_Click()
    initTask ("IMPORT")
End Sub

Private Sub Option7_Click()
    initTask ("UPDATE")
End Sub

Private Sub Option8_Click(Index As Integer)
    initTask ("DELETE")
End Sub

Private Sub Text1_Change()
    Option2 = True
End Sub
Private Sub initTask(taskName As Variant)
    Dim mUtil As New XlUtil
    
    btnID = taskName
    
    Set xlApp = mUtil.getXlApp
    
    Select Case taskName
        Case "CREATE"
            Call toggleFrame(0)
            Call mUtil.initComboBox(xlApp, Combo1, "选择Excel清单")
        Case "IMPORT"
            Call toggleFrame(2)
            Call mUtil.initComboBox(xlApp, Combo3, "选择从Excel清单导入")
            Call mUtil.initComboBox(xlApp, Combo4, "选择Excel清单")
        Case "UPDATE"
            Call toggleFrame(1)
            Call mUtil.initComboBox(xlApp, Combo2, "选择Excel清单")
        Case "DELETE"
            Call toggleFrame(3)
            Call mUtil.initComboBox(xlApp, Combo6, "选择Excel清单")
    End Select
    
    Set mUtil = Nothing
End Sub
Private Sub toggleFrame(showIndex As Integer)
    Dim f As Frame
    For Each f In Frame2
        f.Visible = False
    Next
    Frame2(showIndex).Visible = True
    Frame2(showIndex).Top = 1080
    Frame2(showIndex).Left = 120
   
End Sub

Private Sub handleTask(strID As Variant)
    Dim mUtil As New XlUtil
    
    Me.Hide
    xlApp.Visible = True
    
    Select Case strID
        Case "CREATE"
            Call createSheet
        Case "IMPORT"
            Call importSheets
        Case "UPDATE"
            Call updateSheets
        Case "DELETE"
            Call deleteSheets
            
    End Select
    
    Me.Show 0
    Set mUtil = Nothing
End Sub
Private Sub createSheet()
    Dim mUtil As New XlUtil
    Dim xCreater As New Creater
    Dim mBook As Workbook
    Dim sNum
    
    Set mBook = mUtil.getBook(xlApp, Combo1.Text)
    If Text1.Text <> "" Then
        sNum = CInt(Text1.Text)
    Else
        sNum = mBook.Worksheets.Count
    End If
    
    Call xCreater.addNewSheet(mBook, sNum)
    Call xCreater.formatNewSheet
End Sub
Private Sub updateSheets()
    Dim mUtil As New XlUtil
    Dim mUpdater As New Updater
    Dim mBook As Workbook
    
    Set mBook = mUtil.getBook(xlApp, Combo2.Text)
    Debug.Print mBook.Name
    
    If Option3 = True Then
        Call mUpdater.updateTabs(mBook)
    End If
    
    If Option4 = True Then
        Call mUpdater.updateLinks(mBook)
    End If
    
    Set mUtil = Nothing
    Set mUpdater = Nothing
End Sub
Public Sub importSheets()
    Dim mUtil As New XlUtil
    Dim mImporter As New Importer
    Dim mBook As Workbook, oBook As Workbook
    Dim oStart, oEnd, mAfter, mFirst
    
    Set oBook = mUtil.getBook(xlApp, Combo3.Text)
    Set mBook = mUtil.getBook(xlApp, Combo4.Text)
    
    If Option15(0) = True Then
        oStart = mUtil.getFirstSheetNum(oBook)
        oEnd = oBook.Sheets.Count
    End If
    If Text3.Text <> "" And Text3.Text <> "" Then
        oStart = Text3.Text
        oEnd = Text4.Text
    End If
    
    If Option11(0) = True Then
        mAfter = mBook.Sheets.Count
    End If
    If Text2.Text <> "" Then
        mFirst = mUtil.getFirstSheetNum(mBook)
        mAfter = Text2.Text + mFirst - 1
    End If
    
    Call mImporter.importBook(oBook, mBook, oStart, oEnd, mAfter)
    
    Set mUtil = Nothing
    Set mImporter = Nothing
End Sub
Public Sub deleteSheets()
    Dim mUtil As New XlUtil
    Dim mDeleter As New Deleter
    Dim mBook As Workbook
    Dim oStart, oEnd, mFirst
    
    xlApp.DisplayAlerts = False
    
    Set mBook = mUtil.getBook(xlApp, Combo6.Text)
    
    If Option31(0) = True Then
        oStart = mUtil.getFirstSheetNum(mBook)
        oEnd = mBook.Sheets.Count
    End If
    If Text5.Text <> "" And Text6.Text <> "" Then
        mFirst = mUtil.getFirstSheetNum(mBook)
        oStart = Text6.Text + mFirst - 1
        oEnd = Text5.Text + mFirst - 1
    End If
    
    
    Call mDeleter.deleteBook(mBook, oStart, oEnd)
    
    xlApp.DisplayAlerts = True
    Set mUtil = Nothing
    Set mDeleter = Nothing
End Sub

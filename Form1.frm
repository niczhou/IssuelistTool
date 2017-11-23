VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "问题清单工具"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   8145
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Height          =   2055
      Index           =   1
      Left            =   2280
      TabIndex        =   12
      Top             =   2280
      Width           =   5055
      Begin VB.Frame Frame6 
         Caption         =   "Frame4"
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   4695
         Begin VB.OptionButton Option10 
            Caption         =   "更新汇总清单"
            Height          =   375
            Left            =   3120
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "仅更新标签"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "更新清单链接"
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4695
         Begin VB.ComboBox Combo2 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   4455
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   4695
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2640
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "在"
            Height          =   375
            Left            =   2160
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option1 
            Caption         =   "在最后"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "之后"
            Height          =   255
            Left            =   3240
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4695
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   4455
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择工具"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Option7 
         Caption         =   "更新清单链接"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "导入问题清单"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "创建新问题"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option9 
         Caption         =   "批量输入"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "批量删除"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   120
      Y2              =   4320
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
Private Sub Form_Initialize()
'    frameList = Array(Frame2, Frame5)
End Sub
Private Sub Form_Load()
'    Debug.Print Frame2.Name
    Option7 = True
    Option4 = True
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
    btnID = "IMPORT"
End Sub

Private Sub Option7_Click()
    initTask ("UPDATE")
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
            Call mUtil.initComboBox(xlApp, Combo1)
        Case "IMPORT"
        
        Case "UPDATE"
            Call toggleFrame(1)
            Call mUtil.initComboBox(xlApp, Combo2)
            
    End Select
    
    Set mUtil = Nothing
End Sub
Private Sub toggleFrame(showIndex As Integer)
    Dim f As Frame
    For Each f In Frame2
        f.Visible = False
    Next
    Frame2(showIndex).Visible = True
   
End Sub

Private Sub handleTask(strID As Variant)
    Dim mUtil As New XlUtil
    
    Me.Hide
    xlApp.Visible = True
    
    Select Case strID
        Case "CREATE"
            Call createTask
            
        Case "IMPORT"

        Case "UPDATE"
            Call updateTask
            
    End Select
    
    Me.Show 0
    Set mUtil = Nothing
End Sub
Private Sub createTask()
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
Private Sub updateTask()
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

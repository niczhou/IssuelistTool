VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "问题清单工具"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7545
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   4575
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   2280
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   180
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Left            =   2760
            TabIndex        =   14
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4575
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   240
            TabIndex        =   9
            Text            =   "Combo1"
            Top             =   240
            Width           =   4095
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择工具"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "更新清单"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "导入清单"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "创建新问题"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public btnID
Public xlApp As Excel.Application
Private Sub Command1_Click()
    Call jumpTo("CREATE")
End Sub
Private Sub Command2_Click()
    Call jumpTo("IMPORT")
End Sub

Private Sub Command3_Click()
    Call jumpTo("UPDATE")
End Sub

Private Sub Form_Initialize()

    On Error GoTo ErrorHandler
    
    Set xlApp = GetObject(, "Excel.Application")
    Debug.Print xlApp.Workbooks.Count
    
ErrorHandler:
    Exit Sub
    Set xlApp = CreateObject("Excel.Application")
    Debug.Print xlApp.Workbooks.Count
    
End Sub

Private Sub jumpTo(strID As String)
    Dim xUtil As New XlUtil
 
    btnID = strID
    Me.Hide
    
    Select Case strID
        Case "CREATE"
            Form2.Show 0
            xUtil.initComboBox xlApp, Form2.Combo1
            Form2.Option1 = True
        Case "IMPORT"
            Form3.Show 0
            xUtil.initComboBox xlApp, Form3.Combo1
        Case "UPDATE"
'            Form4.Show 0
'            xUtil.initComboBox xlApp, Form4.Combo1
    End Select
    
    Set xUtil = Nothing
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "问题清单工具"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   6030
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "选择工具"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   3840
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   2160
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "更新清单"
         Height          =   495
         Left            =   3840
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "导入清单"
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "创建新问题"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   5055
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
        Case "IMPORT"
            Form3.Show 0
            xUtil.initComboBox xlApp, Form3.Combo1
        Case "UPDATE"
'            Form4.Show 0
'            xUtil.initComboBox xlApp, Form4.Combo1
    End Select
    
    Set xUtil = Nothing
End Sub

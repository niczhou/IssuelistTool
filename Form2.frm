VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "����������"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "ѡ���������λ��"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "��"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�����"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "ҳ֮��"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ�����Excel�ļ�"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Dim newSheet As Worksheet
    Dim xCreater As New Creater
    Dim xUtil As New XlUtil
    Dim xBook As Workbook

    Form1.xlApp.Visible = True
    Set xBook = xUtil.getBook(Form1.xlApp, Combo1.Text)
    Debug.Print xBook.Name

    If Text1.Text <> "" Then
        Set newSheet = xBook.Sheets.Add(after:=xBook.Sheets(CInt(Text1.Text)))
    ElseIf Option1 = True Then
        Set newSheet = xBook.Sheets.Add(after:=xBook.Sheets(xBook.Sheets.Count))
    End If

    Call xCreater.formatNewSheet(newSheet)
    
    
End Sub

Private Sub Command2_Click()
    Unload Me
    Form1.Show 0
End Sub

'Private Function addNewSheet(xlBook As Workbook, afterSheetNum) As Worksheet
'    Set addNewSheet = xlBook.Sheets.Add(after:=Sheets(afterSheetNum))
'End Function

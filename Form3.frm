VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form3"
   ScaleHeight     =   4320
   ScaleWidth      =   4575
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "在"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "在最后"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "选择导入后汇总清单"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择需导入的问题清单"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form3"
   ScaleHeight     =   5250
   ScaleWidth      =   4575
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

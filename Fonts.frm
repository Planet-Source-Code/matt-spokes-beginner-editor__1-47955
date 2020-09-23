VERSION 5.00
Begin VB.Form frmFonts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Font"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2130
   Icon            =   "Fonts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   2130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox FSize 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox FFace 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Font Size :"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Font Face :"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmMain.HTML.SelRTF = "<font face=" & Chr(34) & FFace.Text & Chr(34) & " size=" & FSize.Text & ">Text</font>"
Unload Me
End Sub

Private Sub Form_Load()
For i = 0 To Screen.FontCount
FFace.AddItem Screen.Fonts(i)
Next i
For s = 1 To 7
FSize.AddItem (s)
Next s
End Sub

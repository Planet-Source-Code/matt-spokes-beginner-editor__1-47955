VERSION 5.00
Begin VB.Form Script 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Script"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "InsertScript.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Code 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Script Language"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.TextBox OT 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton O 
         Caption         =   "Other"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton JS 
         Caption         =   "JavaScript"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton VBS 
         Caption         =   "VBScript"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Code :"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If VBS.Value = True Then
frmMain.HTML.SelRTF = "<script language=" & Chr(34) & "VBScript" & Chr(34) & ">" & vbNewLine & Code.Text & vbNewLine & "</script>" & vbNewLine & "<!-- Script Tags Should Go In Between <head> & </head> -->"
End If
If JS.Value = True Then
frmMain.HTML.SelRTF = "<script language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbNewLine & Code.Text & vbNewLine & "</script>" & vbNewLine & "<!-- Script Tags Should Go In Between <head> & </head> -->"
End If
If O.Value = True Then
frmMain.HTML.SelRTF = "<script language=" & Chr(34) & OT.Text & Chr(34) & ">" & vbNewLine & Code.Text & vbNewLine & "</script>" & vbNewLine & "<!-- Script Tags Should Go In Between <head> & </head> -->"
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

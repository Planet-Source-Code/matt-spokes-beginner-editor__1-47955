VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FileLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert A File Link"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   Icon            =   "FileLink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2760
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Link"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox Check1 
         Caption         =   "Non Underlined"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox LinkT 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox FP 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Link Text :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Link File :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FileLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file1 As String
Private Sub Command1_Click()
CD1.Filter = "All Files|*.*"
CD1.DialogTitle = "Insert A File Link"
CD1.ShowOpen
file1 = CD1.FileTitle
FP.Text = CD1.FileName
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Check1.Value = 0 Then
frmMain.HTML.SelRTF = "<a href=""" + file.Text + """>" + LinkT.Text + "</a>"
Else
frmMain.HTML.SelRTF = "<a href=""" + file1 + """ style=text-decoration:none>" + LinkT.Text + "</a>"
End If
frmMain.List1.AddItem FP.Text
frmMain.List2.AddItem file1
FP.Text = ""
LinkT.Text = ""
Unload Me
End Sub

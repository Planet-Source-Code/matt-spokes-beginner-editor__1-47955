VERSION 5.00
Begin VB.Form Link 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Link"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   Icon            =   "Link.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text Link"
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.TextBox Link 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "http://"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox LinkT 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Link Text"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Non Underlined"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Link URL :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Link Text :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = 0 Then
frmMain.HTML.SelRTF = "<a href=""" + Link.Text + """>" + LinkT.Text + "</a>"
Else
frmMain.HTML.SelRTF = "<a href=""" + Link.Text + """ style=text-decoration:none>" + LinkT.Text + "</a>"
End If
Link.Text = "http://"
LinkT.Text = ""
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

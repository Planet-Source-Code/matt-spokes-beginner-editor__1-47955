VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ImageLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Image Link"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   Icon            =   "ImageLink.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image Link"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.TextBox ILink 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "http://"
         Top             =   480
         Width           =   2535
      End
      Begin VB.ComboBox Border 
         Height          =   315
         ItemData        =   "ImageLink.frx":030A
         Left            =   120
         List            =   "ImageLink.frx":0320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove Picture"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox LinkP 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Link URL :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Border Size :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Link Picture :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
   End
End
Attribute VB_Name = "ImageLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file1 As String
Private Sub Command1_Click()
Bor = Border.Text
der = Mid(Bor, 1, 1)
If Not LinkP.Text = "" Then
frmMain.HTML.SelRTF = "<a href=""" + ILink.Text + """>" + "<img src=" + Chr(34) + file1 + Chr(34) + " border=" + der + "></a>"
frmMain.List1.AddItem LinkP.Text
frmMain.List2.AddItem file1
End If
LinkP.Text = ""
ILink.Text = "http://"
Border.ListIndex = -1
Unload Me
End Sub

Private Sub Command2_Click()
CD1.Filter = "Images Files|*.jpg;*.bmp;*.gif"
CD1.DialogTitle = "Insert Image As A Link"
CD1.ShowOpen
LinkP.Text = CD1.FileName
file1 = CD1.FileTitle
End Sub

Private Sub Command3_Click()
LinkP.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

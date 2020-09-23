VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Image"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   Icon            =   "Image.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Img 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton GetImage 
      Caption         =   "..."
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   420
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Image Source : "
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file1 As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub GetImage_Click()
CD1.Filter = "Images Files|*.jpg;*.bmp;*.gif"
CD1.DialogTitle = "Insert Image"
CD1.ShowOpen
Img.Text = CD1.FileName
file1 = CD1.FileTitle
Pic.Picture = LoadPicture(Img.Text)
End Sub

Private Sub Command2_Click()
frmMain.HTML.SelRTF = "<img src=" + Chr(34) + file1 + Chr(34) + ">"
Pic.Picture = LoadPicture()
frmMain.List1.AddItem Img.Text
frmMain.List2.AddItem file1
Img.Text = ""
Unload Me
End Sub


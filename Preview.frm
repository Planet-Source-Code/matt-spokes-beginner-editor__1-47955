VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Preview 
   Caption         =   "Preview"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "Preview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
m = App.path
WB.Navigate m & "\pre.html"
End Sub

Private Sub Form_Resize()
WB.Height = Preview.Height - 400
WB.Width = Preview.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
m = App.path & "\"
For i = 0 To frmMain.List2.ListCount - 1
Kill m & "\" & frmMain.List2.List(i)
Next i
Kill m & "pre.html"
End Sub

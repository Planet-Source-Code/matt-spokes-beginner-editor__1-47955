VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matts HTML Editor"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7110
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton p 
      Caption         =   "<p>"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton br 
      Caption         =   "<br>"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Right 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      ToolTipText     =   "Right Align"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Middle 
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      ToolTipText     =   "Middle Align"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Left 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      ToolTipText     =   "Left Align"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Font 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "Font"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Underlined 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Underlined"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Italic 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Italic"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Bold 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Bold"
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   3600
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8400
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox HTML 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "HTML Code"
      Top             =   360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10186
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Main.frx":030A
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu load 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu pre 
         Caption         =   "&Preview   "
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu ins 
      Caption         =   "&Insert"
      Begin VB.Menu scr 
         Caption         =   "Insert A Script"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu lin 
         Caption         =   "Insert A Link"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu f 
         Caption         =   "Insert A File Link"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu imlink 
         Caption         =   "Insert An Image Link     "
         Shortcut        =   +{F4}
      End
      Begin VB.Menu im 
         Caption         =   "Insert An Image"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu tab 
         Caption         =   "Insert A Table"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu tim 
         Caption         =   "Insert Time And Date"
         Shortcut        =   +{F7}
      End
   End
   Begin VB.Menu col 
      Caption         =   "&Colors"
      Begin VB.Menu codes 
         Caption         =   "&Get color codes"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file1 As String

Public Function GetPath(path As String) As String
For i = 1 To Len(path)
   If Mid(path, i, 1) = "\" Then x = i
Next i
GetPath = Mid(path, 1, x)
End Function



Private Sub Bold_Click()
HTML.SelRTF = "<b>Bold Text Here</b>"
HTML.SetFocus
End Sub

Private Sub br_Click()
HTML.SelRTF = "<br>"
End Sub

Private Sub codes_Click()
Colors.Show vbModal
End Sub


Private Sub f_Click()
FileLink.Show
End Sub

Private Sub Font_Click()
frmFonts.Show
End Sub

Private Sub im_Click()
frmImage.Show
End Sub

Private Sub imlink_Click()
ImageLink.Show vbModal
End Sub

Private Sub Italic_Click()
HTML.SelRTF = "<i>Italic Text Here</i>"
HTML.SetFocus
End Sub

Private Sub Left_Click()
HTML.SelRTF = "<p align=left>Text</p>"
End Sub

Private Sub lin_Click()
Link.Show vbModal
End Sub

Private Sub load_Click()
Dim filenumber
On Error Resume Next
filenumber = FreeFile
CD1.Filter = "HTML Files|*.html"
CD1.DialogTitle = "Open"
CD1.ShowOpen
Open CD1.FileName For Input As #filenumber
HTML.Text = Input(LOF(filenumber), #filenumber)
Close
End Sub

Private Sub Middle_Click()
HTML.SelRTF = "<p align=center>Text</p>"
End Sub

Private Sub new_Click()
m = MsgBox("You will lose all entered data by starting a new project.", vbOKCancel, "New project")
If m = vbOK Then
HTML.Text = ""
List1.Clear
List2.Clear
NewFile.Show vbModal
End If
End Sub



Private Sub p_Click()
HTML.SelRTF = "<p>"
End Sub

Private Sub pre_Click()
m = App.path & "\"
For i = 0 To List1.ListCount - 1
FileCopy List1.List(i), m & List2.List(i)
Next i
Open m & "pre.html" For Output As #1
Print #1, HTML.Text
Close
Preview.Show
End Sub

Private Sub Right_Click()
HTML.SelRTF = "<p align=right>Text</p>"
End Sub

Private Sub save_Click()
On Error GoTo err:
CD1.CancelError = True
CD1.Filter = "HTML Files|*.html"
CD1.DialogTitle = "Save"
m = GetPath(CD1.FileName)
CD1.ShowSave
Open CD1.FileName For Output As #1
Print #1, HTML.Text
Close #1
For i = 0 To List1.ListCount - 1
FileCopy List1.List(i), m & List2.List(i)
Next i
err:
End Sub

Private Sub scr_Click()
Script.Show vbModal
End Sub

Private Sub tab_Click()
Table.Show vbModal
End Sub

Private Sub tim_Click()
HTML.SelRTF = Date & "  -  " & DatePart("h", Time, vbUseSystemDayOfWeek, vbUseSystem) & ":" & DatePart("n", Time, vbUseSystemDayOfWeek, vbUseSystem)
End Sub

Private Sub Underlined_Click()
HTML.SelRTF = "<u>Underlined Text Here</u>"
HTML.SetFocus
End Sub

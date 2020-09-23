VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form NewFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Project"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "New.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "<br>"
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   0
      Width           =   735
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4455
      Left            =   4080
      TabIndex        =   18
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7858
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"New.frx":030A
   End
   Begin VB.ComboBox FS 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "New.frx":038C
      Left            =   0
      List            =   "New.frx":03A5
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4440
      Width           =   3975
   End
   Begin VB.ComboBox FontFace 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "New.frx":0405
      Left            =   0
      List            =   "New.frx":0407
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   4920
      Width           =   8055
   End
   Begin VB.ComboBox VLink 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "New.frx":0409
      Left            =   0
      List            =   "New.frx":041F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3240
      Width           =   3975
   End
   Begin VB.ComboBox ALink 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "New.frx":044B
      Left            =   0
      List            =   "New.frx":0461
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2640
      Width           =   3975
   End
   Begin VB.ComboBox Text 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "New.frx":048D
      Left            =   0
      List            =   "New.frx":04A3
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1440
      Width           =   3975
   End
   Begin VB.ComboBox Link 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "New.frx":04CF
      Left            =   0
      List            =   "New.frx":04E5
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox Title 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
   Begin VB.ComboBox BC 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "New.frx":0511
      Left            =   0
      List            =   "New.frx":0527
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label9 
      Caption         =   "Font Size :"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Font Name :"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Text :    (Note to go to the next line use <br>)"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "Visited Link Color :"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Active Link Color : "
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Text Color : "
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Link Color :"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Title :"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Back Color :"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "NewFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public A As String
Public V As String
Public L As String
Public B As String
Public T As String
Public C As String
Public Main As String
Public Size As String
Public Face As String
Private Sub Command1_Click()
T = Title.Text
B = BC.Text
C = Text.Text
L = Link.Text
A = ALink.Text
V = VLink.Text
Main = Text1.Text
m = FS.Text
Size = Mid(m, 1, 1)
Face = Chr(34) & FontFace.Text & Chr(34)
frmMain.HTML.Text = "<html>" & vbNewLine & "<head>" & vbNewLine & "<title>" & T & "</title>" & vbNewLine & "</head>"
frmMain.HTML.Text = frmMain.HTML.Text & vbNewLine & "<body bgcolor=" & B & " link=" & L & " alink=" & A & " vlink=" & V & " text=" & C & ">" & vbNewLine
frmMain.HTML.Text = frmMain.HTML.Text & "<font face=" & Face & " size=" & Size & ">" & vbNewLine
frmMain.HTML.Text = frmMain.HTML.Text & Main
frmMain.HTML.Text = frmMain.HTML.Text & vbNewLine & "</font>" & vbNewLine & "</body>" & vbNewLine & "</html>"
Title.Text = ""
BC.ListIndex = -1
Text.ListIndex = -1
Link.ListIndex = -1
ALink.ListIndex = -1
VLink.ListIndex = -1
Text1.Text = ""
FS.ListIndex = -1
FontFace.ListIndex = -1
Unload Me
End Sub

Private Sub Command2_Click()
Text1.SelRTF = "<br>"
Text1.SetFocus
End Sub

Private Sub Form_Load()
matt = ""
For i = 0 To Screen.FontCount
FontFace.AddItem Screen.Fonts(i)
Next i
End Sub

VERSION 5.00
Begin VB.Form Colors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "Colors.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Color 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   3960
      ScaleHeight     =   1995
      ScaleWidth      =   3435
      TabIndex        =   14
      Top             =   0
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   285
      Left            =   3240
      TabIndex        =   13
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Code 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "#000000"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.PictureBox Blue 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   3075
      TabIndex        =   11
      Top             =   1440
      Width           =   3135
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
   Begin VB.HScrollBar BlueScroll 
      Height          =   255
      Left            =   480
      Max             =   255
      TabIndex        =   9
      Top             =   1200
      Width           =   3375
   End
   Begin VB.PictureBox Green 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   840
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.HScrollBar GreenScroll 
      Height          =   255
      Left            =   480
      Max             =   255
      TabIndex        =   6
      Top             =   600
      Width           =   3375
   End
   Begin VB.PictureBox Red 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   3075
      TabIndex        =   5
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar RedScroll 
      Height          =   255
      Left            =   480
      Max             =   255
      TabIndex        =   3
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox B 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "00"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox G 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "00"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox R 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "00"
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function LeftPad(Value, Size As Long, Optional PadCharacter As String = " ") As String
    LeftPad = "" & Value
    While Len(LeftPad) < Size
        LeftPad = PadCharacter & LeftPad
    Wend
End Function

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText Code.Text
Unload Me
End Sub

Private Sub RedScroll_Change()
Red.BackColor = RGB(RedScroll.Value, 0, 0)
Color.BackColor = RGB(RedScroll.Value, GreenScroll.Value, BlueScroll.Value)
R.Text = LeftPad(Hex(RedScroll.Value), 2, 0)
Code.Text = "#" & LeftPad(Hex(RedScroll.Value), 2, 0) & LeftPad(Hex(GreenScroll.Value), 2, 0) & LeftPad(Hex(BlueScroll.Value), 2, 0)
End Sub

Private Sub RedScroll_Scroll()
RedScroll_Change
End Sub
Private Sub GreenScroll_Change()
Green.BackColor = RGB(0, GreenScroll.Value, 0)
Color.BackColor = RGB(RedScroll.Value, GreenScroll.Value, BlueScroll.Value)
G.Text = LeftPad(Hex(GreenScroll.Value), 2, 0)
Code.Text = "#" & LeftPad(Hex(RedScroll.Value), 2, 0) & LeftPad(Hex(GreenScroll.Value), 2, 0) & LeftPad(Hex(BlueScroll.Value), 2, 0)
End Sub

Private Sub GreenScroll_Scroll()
GreenScroll_Change
End Sub

Private Sub BlueScroll_Change()
Blue.BackColor = RGB(0, 0, BlueScroll.Value)
Color.BackColor = RGB(RedScroll.Value, GreenScroll.Value, BlueScroll.Value)
B.Text = LeftPad(Hex(BlueScroll.Value), 2, 0)
Code.Text = "#" & LeftPad(Hex(RedScroll.Value), 2, 0) & LeftPad(Hex(GreenScroll.Value), 2, 0) & LeftPad(Hex(BlueScroll.Value), 2, 0)
End Sub

Private Sub BlueScroll_Scroll()
BlueScroll_Change
End Sub

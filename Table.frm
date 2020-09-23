VERSION 5.00
Begin VB.Form Table 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert A Table"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2310
   Icon            =   "Table.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Row 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "0"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox colum 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox Align 
      Height          =   315
      ItemData        =   "Table.frx":030A
      Left            =   0
      List            =   "Table.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Rows :"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Columns :"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Align :"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If IsNumeric(Row.Text) & IsNumeric(colum.Text) Then
 Call AddTable(colum.Text, Row.Text)
 Else
 MsgBox "Insert a valid number"
 Exit Sub
 End If
 Align.ListIndex = -1
End Sub

Function AddTable(ColumnCount As Long, RowCount As Long) As String
On Error Resume Next
'Add table
Dim tmp
Dim j As Long
Dim k As Long
Dim quote$
quote$ = Chr$(34)
tmp = "<table align=" & Chr(34) & Align.Text & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & ">" & vbCrLf
For j = 1 To RowCount
tmp = tmp & "<tr>" & vbCrLf & "<td> </td>" & vbCrLf
If ColumnCount > 1 Then
For k = 2 To ColumnCount
tmp = tmp & "<td> </td>" & vbCrLf
Next k
End If
tmp = tmp & "</tr>" & vbCrLf
Next j
tmp = tmp & "</table>"
frmMain.HTML.SelRTF = tmp
End Function



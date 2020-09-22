VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Lines..."
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Jump to..."
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label lblLines 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4050
      Width           =   3975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi! i wanna tell you that
'everything of this project
'was created by [[.DarkSouL.]]
' -----------------------
'| MSN: dark_soul@123.cl |
' -----------------------
'
'Thanks for voting...! :P
'Lol!

Private Sub cmdGoto_Click()

    Load frmJumpto
    frmJumpto.Show
    Enabled = False

End Sub

Private Sub cmdRead_Click()

    Load frmReadLines
    frmReadLines.Show
    Enabled = False

End Sub

Private Sub Form_Load()

    Show
    Text1 = ""
    
    Do Until i = 1000
        Caption = "LOADING..."
        i = i + 1
        Text = Text & "Line " & i & vbNewLine
        Text1 = "LOADING... PLEASE WAIT SOME MILISECONDS..."
    Loop

    Text1 = Text

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub Text1_Change()

    Text1_Click

End Sub

Private Sub Text1_Click()

    Caption = "Line " & DetermineLine(Text1) & ", column " & DetermineColumn(Text1) & " - Created by [[.DarkSouL.]] (dark_soul@123.cl)"
    lblLines = "Total lines " & GetTotalLines(Text1)
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    Text1_Click

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

    Text1_Click

End Sub


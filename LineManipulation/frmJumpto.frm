VERSION 5.00
Begin VB.Form frmJumpto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jump To..."
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3150
   Icon            =   "frmJumpto.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtColumn 
      Height          =   285
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdJump 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtLine 
      Height          =   285
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Column:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblLnNum 
      Alignment       =   1  'Right Justify
      Caption         =   "Line:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmJumpto"
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

Private Sub cmdJump_Click()

    On Error GoTo MsgAlert
    Dim LineToJump#, ColumnToJump#
    
    LineToJump = CLng(txtLine)
    ColumnToJump = CLng(txtColumn)

    If LineToJump < 0 Or ColumnToJump < 0 Then GoTo MsgAlert

    frmMain.Enabled = True
    Hide

    SetLineAndColumn frmMain.Text1, LineToJump, ColumnToJump

    Exit Sub
    
MsgAlert:
    MsgBox "You are not alowed to write characters or leave the text boxes in blank. Just write a number upper than 0.", vbExclamation, Caption
    
End Sub

Private Sub txtColumn_GotFocus()

    txtColumn.SelStart = 0
    txtColumn.SelLength = Len(txtLine)

End Sub

Private Sub txtLine_GotFocus()

    txtLine.SelStart = 0
    txtLine.SelLength = Len(txtLine)

End Sub

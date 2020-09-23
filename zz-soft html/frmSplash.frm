VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5865
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   360
      Width           =   6000
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "....Press Any Key"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   0
      Y1              =   3720
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   5880
      X2              =   5880
      Y1              =   3720
      Y2              =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Wise Solution For Website Developers Everywhere!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Width           =   5775
   End
   Begin VB.Line Line2 
      X1              =   5880
      X2              =   5880
      Y1              =   1560
      Y2              =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Mainer.Show
End Sub

Private Sub Frame1_Click()
    Unload Me
    Mainer.Show
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Label5_Click()

End Sub


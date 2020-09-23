VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Insert Picture"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form10"
   ScaleHeight     =   2280
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox c 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Height Of Picture"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox b 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Width Of Picture"
      Top             =   600
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert Picture To Project"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox d 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "picture source here ( i.e. www.hi.com/mypic.bmp )"
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()



End Sub

Private Sub Command2_Click()
Dim a As String
a = """"
Form1.rtb1.SelText = "<img src=" & a & "" & d & "" & a & " width=" & a & "" & b & "" & a & " height=" & a & "" & c & "" & a & ">"
Unload Me
End Sub


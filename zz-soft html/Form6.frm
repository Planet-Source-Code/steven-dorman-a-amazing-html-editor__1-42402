VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Clean Temporary Files"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "All Done"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Files"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"Form6.frx":0442
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Kill App.Path & "\temp000.html"
Kill App.Path & "\temp001.html"
Kill App.Path & "\temp002.html"
Kill App.Path & "\temp005.html"
Kill App.Path & "\tempvws.html"
Kill App.Path & "\temp-date-and time.html"
MsgBox "Temporary Files Sucsessfully Deleted!"
'you want to resume next so if there isnt a certin temp
'file it will go to the next on and try the next
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

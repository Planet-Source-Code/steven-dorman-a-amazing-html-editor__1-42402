VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chnage Website Template"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox text1 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form8.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit Without Saving"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Open App.Path & "\887gh7880.zdll" For Output As #1
Print #1, text1.Text
Close #1
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()

  Open App.Path & "\887gh7880.zdll" For Input As #1
       
      text1.Text = Input$(LOF(1), 1)
   
     Close #1
End Sub

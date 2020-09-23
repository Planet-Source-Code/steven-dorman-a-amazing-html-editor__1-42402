VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Google!"
   ClientHeight    =   1350
   ClientLeft      =   3120
   ClientTop       =   3525
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "SearcH!"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Need Some Info?Qucikly Search Google With This!"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate ("http://www.google.ca/search?q=" & Text1.Text & "&ie=UTF-8&oe=UTF-8&hl=en&meta=")
frmBrowser.WindowState = vbMaximized
Unload Me

End Sub

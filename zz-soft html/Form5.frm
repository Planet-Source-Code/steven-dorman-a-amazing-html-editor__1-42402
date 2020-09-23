VERSION 5.00
Begin VB.Form Form5 
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6915
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   2295
      Left            =   720
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Rcom 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   5280
      Width           =   5775
   End
   Begin VB.Label Label10 
      Caption         =   "Registered Company:"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Rna 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   5040
      Width           =   5535
   End
   Begin VB.Label Label8 
      Caption         =   "Registered to:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label curn 
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   4800
      Width           =   5295
   End
   Begin VB.Label Label6 
      Caption         =   "Current User:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7920
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label5 
      Caption         =   "Need To Contact us? Drop A Line at Seven__dust@hotmail.com Also Please report any found bugs, complements, deathtreats too! "
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7920
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label3 
      Caption         =   "This Is a Nice Little Program Suitable For Any Programmer ( Advanced OR Begginner!)"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Easy Html Editor Is Made By HomeGrown Productions (HgP)"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "ZZ-Soft's Easy Html Editor"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Form5.frx":445C4
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*  *  ******    ******
'*  *  *         *    *
'****  * *****   ******
'*  *  * *   *   *
'*  *  *******   *
'ALL SOURCE IS OWNED BY HOMEGROWN PRODUCTIONS ( HGP )
'PROGRAM IS NOT TO BE EDITED IN ANY WAY , IT IS FOR
'EDUCATIONAL PURPOSES ONLY
'PLEASE VISIT OUR WEBSITE
'HTTP://WWW.H-G-P.TK
'OR DROP A LINE  VIA E-MAIL AT :
'CEO@H-G-P.tk
'(NOTE WEBSITE AND EMAIL ADDRESS INCLUDE NO CAPS)

Private Sub ax_GotFocus()

End Sub

Private Sub Command1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CoolButton1_Click()

End Sub

Private Sub Form_Load()
Form5.Caption = "About HgP'S ZZ-Soft Easy Html Editor -- Current User: " & user & ""
Form5.curn.Caption = "" & user & ""
Form5.Rna.Caption = "" & Ruser & ""
Form5.Rcom.Caption = "" & rcompany & ""
End Sub


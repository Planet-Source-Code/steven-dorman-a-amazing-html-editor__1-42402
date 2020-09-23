VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2190
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Edit Website Template"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save And Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set User Name"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Text Box BackGround Color "
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set BackGround Color"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
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
'seven__dust@hotmail.com
'(NOTE WEBSITE AND EMAIL ADDRESS INCLUDE NO CAPS)

Private Sub Command1_Click()
MsgBox " Current Color is " & background & ""
CommonDialog1.ShowColor
background = CommonDialog1.Color
End Sub

Private Sub Command2_Click()
MsgBox " This only applies to VWS "
MsgBox " Current Color is " & backgroundtext & ""
CommonDialog1.ShowColor
backgroundtext = CommonDialog1.Color
End Sub

Private Sub Command3_Click()
Form7.Show

End Sub

Private Sub Command6_Click()
user = InputBox(" Please enter a name or a nick name to set as the username...")
End Sub

Private Sub Command7_Click()
On Error Resume Next
 Open App.Path & "\settings.zdll" For Output As #1
      Print #1, background
      Print #1, backgroundtext
      Print #1, fonter
       Print #1, user
       Close #1
Call loadset
Call gobar
End Sub


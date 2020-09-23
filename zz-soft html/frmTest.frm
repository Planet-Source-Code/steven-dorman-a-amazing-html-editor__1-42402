VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   6255
      ExtentX         =   11033
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTest.frx":0000
      Left            =   0
      List            =   "frmTest.frx":000D
      TabIndex        =   2
      Text            =   "http://www.microsoft.com"
      Top             =   0
      Width           =   4035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Retrive"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      Height          =   1755
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3600
      Width           =   6255
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo hell
    MousePointer = vbHourglass
    wb.Navigate (Combo1.Text)
    Text2.Text = GetUrlSource(Combo1.Text)
    MousePointer = vbDefault
hell:
    
End Sub


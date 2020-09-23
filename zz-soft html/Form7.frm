VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LinkTopic       =   "Form7"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   4440
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   1680
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label mainl 
      Alignment       =   2  'Center
      Caption         =   "Loading File Viewer........."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label pbl 
      Alignment       =   2  'Center
      Caption         =   "Checking  Versions...."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Width           =   4575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MousePointer = vbHourglass
pb.Value = pb.Value + 2
Dim FileInQuestion2
FileInQuestion2 = Dir(App.Path & "\fileview00.zdll")
If FileInQuestion = "" Then
pb.Value = pb.Value + 2
pb.Value = pb.Value + 2
pb.Value = pb.Value + 2

pb.Value = pb.Value + 2
pbl.Caption = "FileView00.zdll not found , Createing it..."
pb.Value = pb.Value + 2
Open App.Path & "\fileview00.zdll" For Output As #1
Print #1, "password"
Close #1
pb.Value = pb.Value + 2

Else
pb.Value = pb.Value + 2
End If
End Sub


VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form9 
   Caption         =   "Script Editor"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form9"
   ScaleHeight     =   6045
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Script Editor"
      TabPicture(0)   =   "Form10.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scripter"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.CommandButton Command9 
         Caption         =   "Copy To Clipboard"
         Height          =   555
         Left            =   6120
         TabIndex        =   10
         Top             =   4320
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox scripter 
         Height          =   4935
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8705
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form10.frx":001C
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Send To Easy Html"
         Height          =   495
         Left            =   6120
         TabIndex        =   8
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "About This Script Edior"
         Height          =   495
         Left            =   6120
         TabIndex        =   7
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Insert  MsgBox"
         Height          =   495
         Left            =   6120
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Insert Alert"
         Height          =   495
         Left            =   6120
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Insert Sub"
         Height          =   495
         Left            =   6120
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Insert Function"
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open"
         Height          =   495
         Left            =   6120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   495
         Left            =   6120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   6000
         X2              =   6000
         Y1              =   360
         Y2              =   5520
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim filer As String
filer = ""
Mainer.CommonDialog1.Flags = cdlOFNOverwritePrompt
 Mainer.CommonDialog1.Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|VBSCRIPT|*.vbs|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|All Files|*.*| "
Mainer.CommonDialog1.ShowSave
filer = Mainer.CommonDialog1.FileName
If CommonDialog1.FileName <> "" Then

    Open "" & filer & "" For Output As #1
    Print #1, scripter.Text
    Close #1
End If


End Sub

Private Sub Command10_Click()



End Sub

Private Sub Command2_Click()
On Error Resume Next

    With Mainer.CommonDialog1
        .Flags = cdlOFNFileMustExist
        .DialogTitle = "zz-soft Open..."
        .Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|VBSRICPT|*.vbs|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|All Files|*.*| "
        .ShowOpen
                If Len(.FileName) = 0 Then
            Exit Sub
        End If
       
        Open "" & .FileName & "" For Input As #1
      
      scripter.Text = Input$(LOF(1), 1)

     Close #1
    
    End With
    

End Sub

Private Sub Command3_Click()
scripter.SelText = "function newfunction ()"
scripter.SelText = "{"
scripter.SelText = "}"

End Sub

Private Sub Command4_Click()
scripter.SelText = "sub newsub ()"
End Sub

Private Sub Command5_Click()
Dim a As String
a = """"
scripter.SelText = "Alert (" & a & " ALERT TEXT HERE " & a & ")"
End Sub

Private Sub Command6_Click()
Dim a As String
a = """"
scripter.SelText = "msgbox " & a & " message box test here " & a & " "
End Sub

Private Sub Command7_Click()
MsgBox " This is a add on to the html editor made by HomeGrown Productions"
End Sub

Private Sub Command8_Click()

On Error Resume Next
Open App.Path & "\temp005.html" For Output As #1
Print #1, scripter.Text
Close #1

Screen.MousePointer = vbHourglass
            'for setting syntax colors
    m_TextCol = vbYellow
    m_AttribCol = 8388736
    m_TagCol = vbRed
    m_CommentCol = vbGreen
    m_AspCol = 128
      
   
    
    Screen.MousePointer = vbNormal
    pagenumber = pagenumber + 1
    'till the stars , is to create a new form , or actually COPY form1
    Set Mainer.ActiveForm.Form1 = New Form1
    With Form1
        .Caption = DocumentName & " " & pagenumber 'adds the page number
        .Show 'show the NEW form1
        .WindowState = vbMaximized 'mazimize it tofit the mdi form
        '**************
        End With
        'this is for getting the newfile template
         Open App.Path & "\temp005.html" For Input As #1
       Call gobar
       Form1.rtb1.Text = Input$(LOF(1), 1)
     Form1.WindowState = vbMaximized
     Close #1
 Call HtmlHighlight

End Sub

Private Sub Command9_Click()
Clipboard.Clear
Clipboard.SetText (scripter.Text)
MsgBox " Set To Clipboard! "
End Sub

Private Sub Form_Resize()
If WindowState <> vbMinimized Then
WindowState = vbNormal
Else
End If
End Sub

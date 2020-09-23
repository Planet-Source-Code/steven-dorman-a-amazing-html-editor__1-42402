VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form3 
   Caption         =   "Visual Webpage Studio 2.0"
   ClientHeight    =   6435
   ClientLeft      =   1695
   ClientTop       =   1215
   ClientWidth     =   7485
   LinkTopic       =   "Form3"
   ScaleHeight     =   6435
   ScaleWidth      =   7485
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Preview Buffer"
      Height          =   735
      Left            =   4800
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open Existing File"
      DownPicture     =   "Form3.frx":0000
      Height          =   735
      Left            =   3240
      Picture         =   "Form3.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save As..."
      Height          =   735
      Left            =   1680
      Picture         =   "Form3.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New File"
      Height          =   735
      Left            =   0
      Picture         =   "Form3.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox maper 
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form3.frx":1108
      Top             =   4680
      Width           =   11775
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11895
      ExtentX         =   20981
      ExtentY         =   7011
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
      Location        =   "http:///"
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu dghdfh 
         Caption         =   "New"
      End
      Begin VB.Menu saver 
         Caption         =   "Save As"
      End
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu fgnhdfhdfh 
         Caption         =   "-"
      End
      Begin VB.Menu back 
         Caption         =   "Back To ZZ-Soft Easy Html Editor"
         Begin VB.Menu export 
            Caption         =   "Export Current Project To easy html..."
         End
         Begin VB.Menu sss 
            Caption         =   "just Exit VWS"
         End
      End
      Begin VB.Menu fghjfdghj 
         Caption         =   "-"
      End
      Begin VB.Menu jbskjdb 
         Caption         =   "Clean Temporary Files"
      End
      Begin VB.Menu bdjbjdsfbjd 
         Caption         =   "-"
      End
      Begin VB.Menu eevev 
         Caption         =   "exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu sa 
         Caption         =   "Select All"
      End
      Begin VB.Menu copy 
         Caption         =   "copy"
      End
      Begin VB.Menu cut 
         Caption         =   "cut"
      End
      Begin VB.Menu paste 
         Caption         =   "paste"
      End
      Begin VB.Menu vcbmx 
         Caption         =   "-"
      End
      Begin VB.Menu sfasf 
         Caption         =   "Clear Preview Buffer"
      End
      Begin VB.Menu fghkm 
         Caption         =   "-"
      End
      Begin VB.Menu safasf 
         Caption         =   "Clear Clipboard"
      End
   End
   Begin VB.Menu qt 
      Caption         =   "Quick Tags"
      Begin VB.Menu uad 
         Caption         =   "underlined"
      End
      Begin VB.Menu iii 
         Caption         =   "italicized"
      End
      Begin VB.Menu bbb 
         Caption         =   "bold"
      End
      Begin VB.Menu eryj 
         Caption         =   "-"
      End
      Begin VB.Menu eae 
         Caption         =   "break"
      End
      Begin VB.Menu vsbv 
         Caption         =   "horizontal rule"
      End
      Begin VB.Menu kjkj 
         Caption         =   "-"
      End
      Begin VB.Menu lll 
         Caption         =   "link tags"
      End
      Begin VB.Menu hehd 
         Caption         =   "marquee"
      End
      Begin VB.Menu sssr 
         Caption         =   "sentence maker"
      End
   End
   Begin VB.Menu xf 
      Caption         =   "help"
      Begin VB.Menu sasf 
         Caption         =   "Important Information"
      End
   End
End
Attribute VB_Name = "Form3"
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

Private Sub Text2_Change()

End Sub

Private Sub d_Click()

End Sub

Private Sub bbb_Click()
On Error Resume Next
maper.SelText = "<B>    </b>"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Form3.Show
Timer1.Enabled = True
  Open App.Path & "\887gh7880.zdll" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
     Form3.WindowState = vbMaximized
     Close #1
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim filer As String
filer = ""
CommonDialog1.Flags = cdlOFNOverwritePrompt
 CommonDialog1.Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|zz-soft VWS project file|*.zvws|All Files|*.*| "
CommonDialog1.ShowSave
filer = CommonDialog1.FileName
If CommonDialog1.FileName <> "" Then
    Open "" & filer & "" For Output As #1
    Print #1, Form3.maper.Text
    Close #1
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next

    With CommonDialog1
        .Flags = cdlOFNFileMustExist
        .DialogTitle = "zz-soft Open..."
      .Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|zz-soft VWS project file|*.zvws|All Files|*.*| "
        .ShowOpen
                If Len(.FileName) = 0 Then
            Exit Sub
        End If
        Set Form1 = New Form1
        Open "" & .FileName & "" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
       Close #1
    
    End With
    
    Me.Caption = App.Title & " - " & CDlg.FileTitle
    Me.Caption = App.Title & " - " & CDlg.FileName
    
End Sub

Private Sub Command4_Click()
On Error Resume Next
Open App.Path & "\tempvws.html" For Output As #1
Print #1, ""
Close #1


WebBrowser1.Navigate App.Path & "\tempvws.html"


End Sub

Private Sub Command5_Click()

End Sub

Private Sub cut_Click()
On Error Resume Next
 Clipboard.Clear
    Clipboard.SetText Me.maper.SelText, 1
    Me.maper.SelText = ""
End Sub

Private Sub dghdfh_Click()
On Error Resume Next
Form3.Show
Timer1.Enabled = True
  Open App.Path & "\887gh7880.zdll" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
     Form3.WindowState = vbMaximized
     Close #1
End Sub

Private Sub eae_Click()
On Error Resume Next
maper.SelText = "<br>"
End Sub

Private Sub erewa_Click()
If Timer1.Enabled = True Then
Timer1.Interval = "3500"
MsgBox " refresh rate set to 3500 milliseconds "
Else
MsgBox " Auto Refresh Is Not Active! "

End If
End Sub

Private Sub eevev_Click()
Unload Me
End Sub

Private Sub export_Click()
On Error Resume Next
 Screen.MousePointer = vbHourglass
            
    m_TextCol = vbYellow
    m_AttribCol = 8388736
    m_TagCol = vbRed
    m_CommentCol = vbGreen
    m_AspCol = 128
      
   
    
    Screen.MousePointer = vbNormal
 pagenumber = pagenumber + 1
    Set Form1 = New Form1
    With Form1
        .Caption = DocumentName & " " & pagenumber
        .Show
        .WindowState = vbMaximized
        End With
        Open App.Path & "\temp002.html" For Output As #1
      Print #1, Form3.maper.Text
    
     Close #1
     '***
         Open App.Path & "\temp002.html" For Input As #1
       Form1.rtb1.Text = Input$(LOF(1), 1)
     Form1.WindowState = vbMaximized
    Form3.Hide

     Close #1
     Call HtmlHighlight
     
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
'Make sure user really wants to exit
Dim Response As Integer
Response = MsgBox("Are you sure you want to exit Visual Website Studio?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit Editor?")
If Response = vbNo Then
  Cancel = Cancel + 1
  Exit Sub
Else

UnloadMode = UnloadMode + 1
  End If
End Sub

Private Sub gfhj_Click()
If Timer1.Enabled = True Then
Timer1.Interval = "1200"
MsgBox " refresh rate set to 1200 milliseconds "
Else
MsgBox " Auto Refresh Is Not Active! "

End If
End Sub

Private Sub hbnk_Click(Index As Integer)
If Timer1.Enabled = True Then
Timer1.Interval = "1000"
MsgBox " refresh rate set to 1000 milliseconds "
Else
MsgBox " Auto Refresh Is Not Active! "

End If
End Sub

Private Sub hhh_Click()
If Timer1.Enabled = True Then
Timer1.Interval = "750"
MsgBox " refresh rate set to 750 milliseconds "
Else
MsgBox " Auto Refresh Is Not Active! "

End If
End Sub

Private Sub hehd_Click()
On Error Resume Next
maper.SelText = "<marquee bgcolor=  >      </marquee>"
End Sub

Private Sub iii_Click()
On Error Resume Next
maper.SelText = "<i>   </i>"
End Sub

Private Sub iooi_Click()


maper.SelText = "<pre> "




End Sub

Private Sub jbskjdb_Click()
Form6.Show
End Sub

Private Sub lll_Click()
On Error Resume Next
maper.SelText = "<a href=  >   <a/>"
End Sub

Private Sub maper_Change()
On Error Resume Next
Open App.Path & "\tempvws.html" For Output As #1
Print #1, Form3.maper.Text
Close #1


WebBrowser1.Navigate App.Path & "\tempvws.html"
End Sub

Private Sub open_Click()
On Error Resume Next

    With CommonDialog1
        .Flags = cdlOFNFileMustExist
        .DialogTitle = "zz-soft Open..."
      .Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|zz-soft VWS project file|*.zvws|All Files|*.*| "
        .ShowOpen
                If Len(.FileName) = 0 Then
            Exit Sub
        End If
        Set Form1 = New Form1
        Open "" & .FileName & "" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
       Close #1
    
    End With
    
    Me.Caption = App.Title & " - " & CDlg.FileTitle
    Me.Caption = App.Title & " - " & CDlg.FileName
    SB.Panels("Status").Text = "File [" & CDlg.FileTitle & "] are opened"
End Sub

Private Sub paste_Click()
On Error Resume Next
 maper.SelText = Clipboard.GetText(1)
End Sub

Private Sub rsdfh_Click()
If Timer1.Enabled = True Then
Timer1.Interval = "1500"
MsgBox " refresh rate set to 1500 milliseconds "
Else
MsgBox " Auto Refresh Is Not Active! "

End If
End Sub

Private Sub sa_Click()
On Error Resume Next
Me.maper.SelStart = 0
    Me.maper.SelLength = Len(Me.maper.Text)
    Me.maper.SetFocus
End Sub

Private Sub safasf_Click()
Clipboard.Clear

End Sub

Private Sub sasf_Click()
MsgBox " Soon A Patch will be released and it will be version 3.1 "
MsgBox " This patch will mainly be based around VWS , it will upgrade vws to version 2.1 , wich will have many new features"
End Sub

Private Sub saver_Click()
On Error Resume Next
Dim filer As String
filer = ""
CommonDialog1.Flags = cdlOFNOverwritePrompt
 CommonDialog1.Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|zz-soft VWS project file|*.zvws|All Files|*.*| "
CommonDialog1.ShowSave
filer = CommonDialog1.FileName
If CommonDialog1.FileName <> "" Then
    Open "" & filer & "" For Output As #1
    Print #1, Form3.maper.Text
    Close #1
End If
End Sub

Private Sub Slider1_Click()




End Sub

Private Sub sfasf_Click()
On Error Resume Next
Open App.Path & "\tempvws.html" For Output As #1
Print #1, ""
Close #1


WebBrowser1.Navigate App.Path & "\tempvws.html"


End Sub

Private Sub sss_Click()
On Error Resume Next
rep% = MsgBox("Do you want to quit?", vbQuestion + vbYesNo)
If rep% = vbYes Then
Form3.Hide
Timer1.Enabled = False
Else
Exit Sub
End If
End Sub

Private Sub sssr_Click()
On Error Resume Next
insent = InputBox(" What Would You Like The Sentence To Say?")
sizer = InputBox("What Would You Like The Size To BE?")
MsgBox " plese pick a font face "
 CommonDialog1.Flags = cdlCFBoth Or cdlCFTTOnly Or cldCFEffects
    CommonDialog1.ShowFont
  
        typer = CommonDialog1.FontName
       
    MsgBox " ignoring all options but font name... "
    repr% = MsgBox("Bold?", vbQuestion + vbYesNo)
If repr% = vbYes Then
b = "<b>"
be = "</b>"
Else
b = ""
be = ""
End If
reprr% = MsgBox("italic?", vbQuestion + vbYesNo)
If reprr% = vbYes Then
i = "<i>"
ie = "</i>"
Else
i = ""
ie = ""
End If
reprrrr% = MsgBox("underlined?", vbQuestion + vbYesNo)
If reprrrr% = vbYes Then
U = "<u>"
ue = "</u>"
Else
U = ""
ue = ""
End If
colo = InputBox(" What color would you like it?")
MsgBox " compileing... "

maper.SelText = "" & b & "" & U & "" & i & "<font color=" & colo & " size=" & sizer & " face=" & typer & " >" & insent & "</font> " & ue & " " & be & " " & ie & " "
End Sub

Private Sub tdfg_Click()
If Timer1.Enabled = True Then
Timer1.Interval = "2000"
MsgBox " refresh rate set to 2000 milliseconds "
Else
MsgBox " Auto Refresh Is Not Active! "

End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Open App.Path & "\tempvws.html" For Output As #1
Print #1, Form3.maper.Text
Close #1


WebBrowser1.Navigate App.Path & "\tempvws.html"

End Sub

Private Sub uad_Click()
On Error Resume Next
maper.SelText = "<u>   </u>"
End Sub

Private Sub vsbv_Click()
On Error Resume Next
maper.SelText = "<hr size=  color=  >"
End Sub

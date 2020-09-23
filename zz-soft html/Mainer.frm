VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Mainer 
   BackColor       =   &H00000040&
   Caption         =   "ZZ-Soft Easy Html Editor 3.0 --"
   ClientHeight    =   5715
   ClientLeft      =   1695
   ClientTop       =   1425
   ClientWidth     =   9645
   Icon            =   "Mainer.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ProgressBar Pb 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5460
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "ZZ-Soft Easy Html Editor 3.0"
            TextSave        =   "ZZ-Soft Easy Html Editor 3.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1/12/03"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:51 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu new 
         Caption         =   "New Web Document"
         Shortcut        =   ^N
      End
      Begin VB.Menu dfgsfdyhaddfj 
         Caption         =   "New Blank Page"
         Shortcut        =   ^B
      End
      Begin VB.Menu dtryjdghjdghjd 
         Caption         =   "-"
      End
      Begin VB.Menu open 
         Caption         =   "Open File"
      End
      Begin VB.Menu saveasa 
         Caption         =   "Save As ..."
      End
      Begin VB.Menu dfgffffgjf 
         Caption         =   "-"
      End
      Begin VB.Menu dgd 
         Caption         =   "Open Visual WebPage Studio 2.0"
         Begin VB.Menu ffgf 
            Caption         =   "...transfer current"
         End
         Begin VB.Menu ddbd 
            Caption         =   "New VWS Project"
         End
      End
      Begin VB.Menu nbngfggfgf 
         Caption         =   "-"
      End
      Begin VB.Menu clean 
         Caption         =   "Clean Temporary Files"
         Shortcut        =   ^E
      End
      Begin VB.Menu fgjgfjfdgj 
         Caption         =   "-"
      End
      Begin VB.Menu exite 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu editr 
      Caption         =   "Edit"
      Begin VB.Menu sa 
         Caption         =   "Select All "
         Shortcut        =   ^A
      End
      Begin VB.Menu lop 
         Caption         =   "cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu cuop 
         Caption         =   "copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu fdsd 
         Caption         =   "paste"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sgfhxfgj 
         Caption         =   "-"
      End
      Begin VB.Menu ewew 
         Caption         =   "PREVIEW"
      End
      Begin VB.Menu uhuygyt 
         Caption         =   "-"
      End
      Begin VB.Menu csss 
         Caption         =   "Font Selector Selector( font name)"
         Shortcut        =   ^G
      End
      Begin VB.Menu set 
         Caption         =   "Settings"
      End
      Begin VB.Menu linea 
         Caption         =   "-"
      End
      Begin VB.Menu dss 
         Caption         =   "refresh syntax"
         Shortcut        =   {F12}
      End
      Begin VB.Menu sfasfasfasfasfasfa 
         Caption         =   "Clear Clipboard"
      End
      Begin VB.Menu jrfj6jfjfgj 
         Caption         =   "Insert To Clipboard..."
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu snip 
      Caption         =   "Snippets..."
      Begin VB.Menu hsnip 
         Caption         =   "Snippets"
      End
   End
   Begin VB.Menu qt 
      Caption         =   "Quick Tags"
      Begin VB.Menu it 
         Caption         =   "italic tags"
      End
      Begin VB.Menu btq 
         Caption         =   "Bold Tags"
      End
      Begin VB.Menu ut 
         Caption         =   "underlined tags"
      End
      Begin VB.Menu qqqqq 
         Caption         =   "-"
      End
      Begin VB.Menu bt 
         Caption         =   "Break"
      End
      Begin VB.Menu ht 
         Caption         =   "horizontal rule"
      End
      Begin VB.Menu qqqqqqqqqqqqqqq 
         Caption         =   "-"
      End
      Begin VB.Menu lt 
         Caption         =   "link tags"
         Shortcut        =   ^L
      End
      Begin VB.Menu sadgs 
         Caption         =   "Marquee"
         Shortcut        =   ^M
      End
      Begin VB.Menu table 
         Caption         =   "Table"
         Shortcut        =   ^T
      End
      Begin VB.Menu dfhdfhdfhdfhdfhdfhdfh 
         Caption         =   "-"
      End
      Begin VB.Menu sdgvas 
         Caption         =   "Button"
      End
      Begin VB.Menu qawsdfa 
         Caption         =   "text feild"
      End
      Begin VB.Menu sdgsdgsdgsdgsdgsdgfsd 
         Caption         =   "password field"
      End
      Begin VB.Menu fh 
         Caption         =   "-"
      End
      Begin VB.Menu sf 
         Caption         =   "Sentence Genorator"
      End
      Begin VB.Menu dddddd 
         Caption         =   "section maker"
      End
      Begin VB.Menu DFHDFHDFHDFH 
         Caption         =   "Insert Picture"
      End
   End
   Begin VB.Menu pv 
      Caption         =   "PREVIEW"
   End
   Begin VB.Menu asasae 
      Caption         =   "Other Programs"
      Begin VB.Menu ljkskdfhgshdgbsdgbskjdghs 
         Caption         =   "You Have Been Working Since..."
      End
      Begin VB.Menu dndfhk 
         Caption         =   "Source Retreiver"
      End
      Begin VB.Menu vwsers 
         Caption         =   "Visual Website Studio"
         Begin VB.Menu wdfadgfwrgvscw 
            Caption         =   "...transfer current"
         End
         Begin VB.Menu sssdfs 
            Caption         =   "new VWS projects"
         End
      End
      Begin VB.Menu ssdfs 
         Caption         =   "Script Editor"
         Shortcut        =   ^S
      End
      Begin VB.Menu ghdfbdfhd 
         Caption         =   "Quick Google Search"
      End
   End
   Begin VB.Menu hwpp 
      Caption         =   "Help"
      Begin VB.Menu csv 
         Caption         =   "About"
      End
      Begin VB.Menu hcswe 
         Caption         =   "Help Contents"
      End
      Begin VB.Menu jdrfjkdiofghdghidgfhdfhdfjhdfhdfh 
         Caption         =   "Tips And Tricks"
      End
      Begin VB.Menu dhksdgf 
         Caption         =   "Expected in next verison..."
      End
      Begin VB.Menu awdrgha 
         Caption         =   "dedications , and other info"
      End
   End
End
Attribute VB_Name = "Mainer"
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




Private Sub awdrgha_Click()
On Error Resume Next

Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\dedicationsetc.txt"
frmBrowser.WindowState = vbMaximized
End Sub

Private Sub bt_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<br>" ' inserts tag
Call HtmlHighlight
End Sub

Private Sub btq_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<b>   </b>" ' inserts tag
Call HtmlHighlight
End Sub

Private Sub clean_Click()
Form6.Show
End Sub

Private Sub csss_Click()
On Error Resume Next
   CommonDialog1.Flags = cdlCFBoth Or cdlCFTTOnly Or cldCFEffects
    CommonDialog1.ShowFont
    With Mainer.ActiveForm.rtb1
        .SelText = CommonDialog1.FontName ' inserts font name
       End With
    MsgBox " ignoring all options but font name... "
    Call HtmlHighlight
End Sub

Private Sub csv_Click()
On Error Resume Next
Form5.Show ' for the about
End Sub

Private Sub cuop_Click()
On Error Resume Next
 Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.rtb1.SelText, 1
     
End Sub

Private Sub ddbd_Click()
On Error Resume Next

Form3.Show
Timer1.Enabled = True
  Open App.Path & "\887gh7880.zdll" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
     Form3.WindowState = vbMaximized
     Close #1
     
End Sub

Private Sub dfgsdfhfhcdfhd_Click()
  
End Sub

Private Sub dddddd_Click()
MsgBox " This uses a table to make a section "
Dim a, b, c, d, e As String
a = InputBox("Please Select A Width..")
b = InputBox("Please Select A Height")
c = InputBox("Please Select A Background Color For The Section..")
d = InputBox("Please Select A Border Size..")
e = InputBox("Please Select A Boreder Color...")
MsgBox " Compileing.... "
Call gobar
Mainer.ActiveForm.rtb1.SelText = "<table width=" & a & " height=" & b & " bgcolor=" & c & " border=" & d & " bordercolor=" & e & "><td> <!-- section contents here --> </td></table> "

Call HtmlHighlight

End Sub

Private Sub dfgsfdyhaddfj_Click()
Screen.MousePointer = vbHourglass
            'for setting syntax colors
    m_TextCol = vbYellow
    m_AttribCol = 8388736
    m_TagCol = vbRed
    m_CommentCol = vbGreen
    m_AspCol = 128
      
   
    
    
    pagenumber = pagenumber + 1
    'till the stars , is to create a new form , or actually COPY form1
    Set Form1 = New Form1
    With Form1
        .Caption = DocumentName & " " & pagenumber 'adds the page number
        .Show 'show the NEW form1
        .WindowState = vbMaximized 'mazimize it tofit the mdi form
        '**************
        End With
       
               Call gobar
           Form1.WindowState = vbMaximized
     Close #1
 Call HtmlHighlight
 Screen.MousePointer = vbNormal
End Sub

Private Sub DFHDFHDFHDFH_Click()
Form10.Show
End Sub

Private Sub dhksdgf_Click()
On Error Resume Next

Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\in next version.txt"
frmBrowser.WindowState = vbMaximized
End Sub

Private Sub dndfhk_Click()
frmTest.Show
End Sub

Private Sub dss_Click()
On Error GoTo hell
Call HtmlHighlight
hell:
End Sub

Private Sub ewew_Click()
'for previewing
On Error Resume Next
Open App.Path & "\temp000.html" For Output As #1
Print #1, Mainer.ActiveForm.rtb1.Text
Close #1
Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\temp000.html"
frmBrowser.WindowState = vbMaximized

End Sub

Private Sub exite_Click()
rep% = MsgBox("Do you want to quit?", vbQuestion + vbYesNo)
If rep% = vbYes Then
End
Else
Exit Sub
End If
End Sub

Private Sub fdsd_Click()
On Error Resume Next
 Me.ActiveForm.rtb1.SelText = Clipboard.GetText(1)
 Call HtmlHighlight
    End Sub

Private Sub ffgf_Click()
On Error Resume Next
Form3.Show
Timer1.Enabled = True
 Open App.Path & "\temp001.html" For Output As #1
      Print #1, Mainer.ActiveForm.rtb1.Text
      Close #1
  Open App.Path & "\temp001.html" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
     Form3.WindowState = vbMaximized
     Close #1
End Sub

Private Sub ghdfbdfhd_Click()
Form8.Show
End Sub

Private Sub hcswe_Click()
On Error Resume Next

Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\help.txt"
frmBrowser.WindowState = vbMaximized

End Sub

Private Sub hsnip_Click()
On Error Resume Next
Form2.Show
End Sub

Private Sub ht_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<hr size=   color=  >"
Call HtmlHighlight
End Sub

Private Sub it_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<i>   </i>"
Call HtmlHighlight
End Sub

Private Sub jkns_Click()
Form7.Show
End Sub

Private Sub jdrfjkdiofghdghidgfhdfhdfjhdfhdfh_Click()
On Error Resume Next

Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\tipsandtricks.html"
frmBrowser.WindowState = vbMaximized
End Sub

Private Sub jrfj6jfjfgj_Click()
Dim thetext As String
thetext = InputBox("Please Type The Text You Wish To Add To The ClipBoard")
Clipboard.Clear
Clipboard.SetText (thetext)

End Sub

Private Sub ljkskdfhgshdgbsdgbskjdghs_Click()
MsgBox " You Have been working Since " & timei & ""
End Sub

Private Sub lop_Click()
On Error Resume Next
 Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.rtb1.SelText, 1
    Me.ActiveForm.rtb1.SelText = ""
End Sub

Private Sub lt_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<a href=   >  </a>"
Call HtmlHighlight
End Sub

Private Sub MDIForm_Load()
timei = "" & Time & ""

Mainer.WindowState = vbMaximized
Dim FileInQuestion
FileInQuestion = Dir(App.Path & "\regest.zdll")
If FileInQuestion = "" Then
MsgBox "Welcome New ZZ-Soft Easy Html Editor User!"
MsgBox " To Start Using This Program  , You Must Regester , and this is free."
MsgBox "regestering does not distribute anything on the internet."
MsgBox " Your Info Is Safe With Us! Certified! "
Ruser = InputBox("What Is The Current Computers owners name?")
rcompany = InputBox("what Is His or Her Company?")
MsgBox " thank you for your time! "
Open App.Path & "\regest.zdll" For Output As #1
        On Error Resume Next
      Print #1, Ruser
      Print #1, rcompany
        Close #1

MsgBox " Also : Because this is your first use , no settings have been set."
MsgBox " goto edit then go to settings to set  user name  , and change the programs color settings."
MsgBox " Note : the user name does not do anything , it just personalizes the program further "

Else
Open App.Path & "\regest.zdll" For Input As #1
        On Error Resume Next
      Input #1, Ruser
      Input #1, rcompany
        Close #1
End If
 
'****
Call loadset
Open App.Path & "\temp-date-and time.html" For Output As #1
Print #1, "This Program Was Last Opened On " & Date & ", " & Time & " <BR> this was an auto genorated temporary file made by easy html editor<BR>"
Close #1
Me.Caption = "" & user & "'s ZZ-Soft Easy Html Editor  -- "
new_Click



  
           
                
                End Sub
            

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim a%
    If FormsCount > 0 Then
        Do Until FormsCount = 0
            If Cancelled = True Then
                Cancelled = False
                Cancel = 1
                Exit Sub
            Else
                Unload Me.ActiveForm
            End If
        Loop
    Else
        a = MsgBox("Are You Sure You Want To Exit?", vbYesNo + vbInformation, "ZZ-Softs Easy Html Editor")
        Select Case a
        Case 6
            Unload Me
            End
        Case 7
            Cancel = 1
        End Select
    End If
End Sub


Sub Down(mee As Form)  'Place in form resize
    If mee.WindowState = vbMinimized Then mee.Hide
End Sub


Private Sub MDIForm_Resize()
If Mainer.WindowState = vbMinimized Then
'stay minimized
Else
'maximized if form is resized
Mainer.WindowState = vbMaximized
End If
End Sub


Private Sub new_Click()
On Error Resume Next


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
    Set Form1 = New Form1
    With Form1
        .Caption = DocumentName & " " & pagenumber 'adds the page number
        .Show 'show the NEW form1
        .WindowState = vbMaximized 'mazimize it tofit the mdi form
        '**************
        End With
        'this is for getting the newfile template
         Open App.Path & "\887gh7880.zdll" For Input As #1
       Call gobar
       Form1.rtb1.Text = Input$(LOF(1), 1)
     Form1.WindowState = vbMaximized
     Close #1
 Call HtmlHighlight
End Sub

'simple opening file procedure
Private Sub open_Click()
On Error Resume Next

    With CommonDialog1
        .Flags = cdlOFNFileMustExist
        .DialogTitle = "zz-soft Open..."
        .Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|All Files|*.*| "
        .ShowOpen
                If Len(.FileName) = 0 Then
            Exit Sub
        End If
        Set Form1 = New Form1
        Open "" & .FileName & "" For Input As #1
        Call gobar
       Form1.rtb1.Text = Input$(LOF(1), 1)
     Form1.WindowState = vbMaximized
     Close #1
    
    End With
    
    Me.Caption = App.Title & " - " & CDlg.FileTitle
    Me.Caption = App.Title & " - " & CDlg.FileName
    SB.Panels("Status").Text = "File [" & CDlg.FileTitle & "] are opened"
        HtmlHighlight
  
End Sub


Private Sub pv_Click()
On Error Resume Next
Open App.Path & "\temp000.html" For Output As #1
Print #1, Mainer.ActiveForm.rtb1.Text
Close #1
Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\temp000.html"
frmBrowser.WindowState = vbMaximized



End Sub

Private Sub qawsdfa_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<input type=text value='Your Text Here' name=text>"
Call HtmlHighlight
End Sub

Private Sub sa_Click()
On Error Resume Next
Me.ActiveForm.rtb1.SelStart = 0
    Me.ActiveForm.rtb1.SelLength = Len(Me.ActiveForm.rtb1.Text)
    Me.ActiveForm.rtb1.SetFocus
End Sub

Private Sub sadgs_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<marquee bgcolor= >   </marquee>"
Call HtmlHighlight
End Sub

Private Sub saveasa_Click()
On Error Resume Next
Dim filer As String
filer = ""
CommonDialog1.Flags = cdlOFNOverwritePrompt
 CommonDialog1.Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|All Files|*.*| "
CommonDialog1.ShowSave
filer = CommonDialog1.FileName
If CommonDialog1.FileName <> "" Then
Call gobar
    Open "" & filer & "" For Output As #1
    Print #1, Mainer.ActiveForm.rtb1.Text
    Close #1
End If
End Sub

Private Sub tete_Click()
Form3.Show
End Sub

Private Sub sdgsdgsdgsdgsdgsdgfsd_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<input type=password value= name=password>"
Call HtmlHighlight
End Sub

Private Sub sdgvas_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<input type=button value=button name=button>"
Call HtmlHighlight
End Sub

Private Sub set_Click()
On Error Resume Next
Form4.Show
End Sub

Private Sub sf_Click()
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

Mainer.ActiveForm.rtb1.SelText = "" & b & "" & U & "" & i & "<font color=" & colo & " size=" & sizer & " face=" & typer & " >" & insent & "</font> " & ue & " " & be & " " & ie & " "
Call gobar
Call HtmlHighlight

End Sub

Private Sub sfasfasfasfasfasfa_Click()
Clipboard.Clear
End Sub

Private Sub sfsgsdbdfjhdgndddd_Click()



End Sub

Private Sub ssdfs_Click()
Load Form9
Form9.Show
End Sub

Private Sub sssdfs_Click()
On Error Resume Next

Form3.Show
Timer1.Enabled = True
  Open App.Path & "\887gh7880.zdll" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
     Form3.WindowState = vbMaximized
     Close #1
End Sub

Private Sub table_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<table height= width=>  </table>"
Call HtmlHighlight
End Sub

Private Sub Timer1_Timer()
timei = timei + 1

End Sub

Private Sub ut_Click()
On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<u>   </u>"
Call HtmlHighlight
End Sub

Private Sub www_Click()

End Sub

Private Sub wdfadgfwrgvscw_Click()
On Error Resume Next
Form3.Show
Timer1.Enabled = True
 Open App.Path & "\temp001.html" For Output As #1
      Print #1, Mainer.ActiveForm.rtb1.Text
      Close #1
  Open App.Path & "\temp001.html" For Input As #1
       Form3.maper.Text = Input$(LOF(1), 1)
     Form3.WindowState = vbMaximized
     Close #1
End Sub

Private Sub wefhddfhdfh_Click()

End Sub

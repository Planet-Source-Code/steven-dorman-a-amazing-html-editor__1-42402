VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   3060
   ClientTop       =   3180
   ClientWidth     =   4245
   ForeColor       =   &H000000C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   4245
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   1191
      ButtonWidth     =   609
      ButtonHeight    =   1032
      AllowCustomize  =   0   'False
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Line Style"
            Object.ToolTipText     =   "Horizontal Rule"
            ImageKey        =   "Line Style"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Justify"
            Object.ToolTipText     =   "Line Break"
            ImageKey        =   "Justify"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sum"
            Object.ToolTipText     =   "Sentance Genorator"
            ImageKey        =   "Sum"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "p"
            Key             =   "p"
            Object.ToolTipText     =   "Preview"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   720
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0554
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0666
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0778
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":088A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":099C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AAE
            Key             =   "Line Style"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BC0
            Key             =   "Justify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CD2
            Key             =   "Sum"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   12632064
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0DE4
   End
End
Attribute VB_Name = "Form1"
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
    

Private Declare Function SendMessageLong Lib "User32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1

Public OpenFilename As String


Const WM_COPY = &H301
Const WM_CUT = &H300
Const WM_CLEAR = &H303
Const WM_PASTE = &H302

Public trapUndo As Boolean
Private UndoStack As New Collection
Private RedoStack As New Collection


Public CtlKey As Boolean

Private Sub RichTextBox1_Change()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Save"
           On Error Resume Next
Dim filer As String
filer = ""
Mainer.CommonDialog1.Flags = cdlOFNOverwritePrompt
 Mainer.CommonDialog1.Filter = "Html Web Page|*.html; *.htm| Xml web page|*.xml; *.Xsl; *.xsd|Cgi Page|*.cgi|Active server page|*.asp; *.aspx|Shtml Page|*.shtml; *.shtm|Dhtml Page|*.dhtml|Java Scripts|*.js; *.jav; *.java|Perl Scripts|*.pl; *.pm|PHP Files|*.php; *.php3|SQL Files|*.sql|ZZ-Soft Files|*.zdll|All Files|*.*| "
Mainer.CommonDialog1.ShowSave
filer = Mainer.CommonDialog1.FileName
If Mainer.CommonDialog1.FileName <> "" Then
Call gobar
    Open "" & filer & "" For Output As #1
    Print #1, Mainer.ActiveForm.rtb1.Text
    Close #1
End If
        Case "Open"
          On Error Resume Next
 Screen.MousePointer = vbHourglass
            
    m_TextCol = vbYellow
    m_AttribCol = 8388736
    m_TagCol = vbRed
    m_CommentCol = vbGreen
    m_AspCol = 128
      
   
    
    Screen.MousePointer = vbNormal
    With Mainer.CommonDialog1
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
    Call HtmlHighlight
    End With
    
    Me.Caption = App.Title & " - " & CDlg.FileTitle
    Me.Caption = App.Title & " - " & CDlg.FileName
    SB.Panels("Status").Text = "File [" & CDlg.FileTitle & "] are opened"
        Case "New"
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
         Open App.Path & "\887gh7880.zdll" For Input As #1
       Call gobar
       Form1.rtb1.Text = Input$(LOF(1), 1)
     Form1.WindowState = vbMaximized
     Close #1
     Call HtmlHighlight
     
        Case "Bold"
         On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<b>   </b>"
Call HtmlHighlight
        Case "Italic"
           On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<i>   </i>"
Call HtmlHighlight
        Case "Underline"
            On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<u>   </u>"
Call HtmlHighlight
        Case "Line Style"
            On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<HR Size= color= >"
Call HtmlHighlight
        Case "Justify"
           On Error Resume Next
Mainer.ActiveForm.rtb1.SelText = "<BR>"
Call HtmlHighlight
        Case "Sum"
          
'**** ;;;:::: BELOW IS (C) HgP! ::::;;;; ****
'**** YOU MAY USE THIS IS ANY APPLICATION,ASLONG AS YOU
'*** SAY IT WAS MADE BY STEVEN DORMAN!

On Error Resume Next


insent = InputBox(" What Would You Like The Sentence To Say?")
sizer = InputBox("What Would You Like The Size To BE?")
MsgBox " plese pick a font face "
 Mainer.CommonDialog1.Flags = cdlCFBoth Or cdlCFTTOnly Or cldCFEffects
    Mainer.CommonDialog1.ShowFont
  
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
Call HtmlHighlight

Case "p"
On Error Resume Next
Open App.Path & "\temp000.html" For Output As #1
Print #1, Mainer.ActiveForm.rtb1.Text
Close #1
Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\temp000.html"
frmBrowser.WindowState = vbMaximized
    End Select
End Sub





Private Sub Form_Load()
On Error Resume Next
 
  
    
    
 Open App.Path & "\settings.zdll" For Input As #1
 
 
Dim na As String
       Line Input #1, background
       Line Input #1, na
       Line Input #1, backgroundtext
       Close #1

  Screen.MousePointer = vbHourglass
            
    m_TextCol = vbYellow
    m_AttribCol = 8388736
    m_TagCol = vbRed
    m_CommentCol = vbGreen
    m_AspCol = 128
      
   rtb1.Visible = False
rtb1.AutoVerbMenu = True
    rtb1.HideSelection = True
     
   
       
    Call HtmlHighlight
       
   
  rtb1.Visible = True
    rtb1.TabStop = True
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Make sure user really wants to exit
Dim Response As Integer
Response = MsgBox("Are you sure you want to close this project?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit Editor?")
If Response = vbNo Then
  Cancel = Cancel + 1
  Exit Sub
Else
UnloadMode = UnloadMode + 1
  End If
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMaximized Then
      Mainer.Caption = "" & user & "'s ZZ-Soft Easy Html Editor"
    Else
       Mainer.Caption = "" & user & "'s ZZ-Soft Easy Html Editor"
    End If
    If Me.WindowState <> vbMinimized Then
        If Mainer.Visible = True Then
            If Mainer.Visible = True Then
                rtb1.Move 0, 780, Me.ScaleWidth, Me.ScaleHeight
            Else
                rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            End If
        Else
            rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        End If
    End If
End Sub

Private Sub rtb1_Change()
  If Not trapUndo Then Exit Sub


    Dim c%, l&

    
    For c% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c%

    
    newElement.SelStart = RichTxtBox.SelStart
    newElement.TextLen = Len(RichTxtBox.Text)
    newElement.Text = RichTxtBox.Text



End Sub

Private Sub rtb1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    KeyAscii = KeyPressEvent(KeyAscii)
End Sub

Private Sub rtb1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TypedIn As String
    If Shift And vbCtrlMask Then
        If KeyCode > vbKey0 And KeyCode < vbKey7 Then
            Dim HeadingTag As String
            HeadingTag = "<H" & CStr(KeyCode - vbKey0) & "></H" & CStr(KeyCode - vbKey0) & ">"
            InsertTag HeadingTag, True
            PlaceCursor.HeadingTag , 5
            rtb1.SelColor = vbBlack
        Else
            Select Case KeyCode
            Case vbKeyV
               
                Dim a$, S As Long
                S = rtb1SelStart
                a = Clipboard.GetText(vbCFText)
                rtb1.SelText = ""
                rtb1.SelText = a
                HtmlColorCode S, rtb1.SelStart
                
                KeyCode = 0
            Case vbKeyReturn
                InsertTag "<P>", True
                rtb1.SelColor = vbRed
                KeyCode = 0
            Case vbKeySpace
                rtb1.SelColor = vbRed
               rtb1.SelText = "&nbsp;"
                KeyCode = 0
            End Select
        End If
    ElseIf Shift And vbShiftMask Then
        If KeyCode = vbKeyReturn Then
            InsertTag "<BR>", True
            rtb1.SelColor = vbRed
            KeyCode = 0
        End If
    End If
    IsOutsideTag
End Sub

Private Sub RichTxtBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsOutsideTag

End Sub

Public Sub GetEditStatus()
   Dim lLine As Long, lCol As Long
   Dim cCol As Long, lChar As Long, i As Long

   lChar = rtb1.SelStart + 1

   
   lLine = 1 + SendMessageLong(rtb1.hWnd, EM_LINEFROMCHAR, _
           rtb1.SelStart, 0&)

   cCol = SendMessageLong(rtb1.hWnd, EM_LINELENGTH, lChar - 1, 0&)

   i = SendMessageLong(rtb1.hWnd, EM_LINEINDEX, lLine - 1, 0&)
   lCol = lChar - i


   sbStatusBar.Panels(1).Text = "Line: " & lLine & ", Col: " & lCol

End Sub


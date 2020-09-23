VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snippets"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6525
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command10 
      Caption         =   "8"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "DropDown Navigation"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Popup Window"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "8"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Highlighted Links"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "8"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Redirection"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      Height          =   3960
      ItemData        =   "Form2.frx":030A
      Left            =   2760
      List            =   "Form2.frx":030C
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Html Tags ---------------------->"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "This is The Snippets Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form2"
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
List2.Clear
List2.AddItem "These are Some Basic Tags For Html"
List2.AddItem "<pre> </pre>"
List2.AddItem "<tt> </tt>"
List2.AddItem "<iframe src=  > </iframe>"
List2.AddItem "<b>  </b>"
List2.AddItem "<i>  </i>"
List2.AddItem "<u>  </u>"
List2.AddItem "<font color=  size= >  </font>"
List2.AddItem "<a href=  >  </a>"
List2.AddItem "<input type= value= onclick=  >"
List2.AddItem "<input type= value= onmouseover=  >"
List2.AddItem "<input type= value= >"
List2.AddItem "<br>"
List2.AddItem "<hr>"
List2.AddItem "<hr size= >"
List2.AddItem "<hr size=  color=  >"
List2.AddItem "<h1>   </h1>"
List2.AddItem "<h2>    </h2>"
List2.AddItem "<h3>     <h3>"
List2.AddItem "<marquee bgcolor=  >      </marquee>"
End Sub

Private Sub List1_Click()
List2.Clear
List2.AddItem "These are Some Basic Tags For Html"
List2.AddItem "<pre> </pre>"
List2.AddItem "<tt> </tt>"
List2.AddItem "<iframe src=  > </iframe>"
List2.AddItem "<b>  </b>"
List2.AddItem "<i>  </i>"
List2.AddItem "<u>  </u>"
List2.AddItem "<font color=  size= >  </font>"
List2.AddItem "<a href=  >  </a>"
List2.AddItem "<input type= value= onclick=  >"
List2.AddItem "<input type= value= onmouseover=  >"
List2.AddItem "<input type= value= >"
List2.AddItem "<br>"
List2.AddItem "<hr>"
List2.AddItem "<hr size= >"
List2.AddItem "<hr size=  color=  >"
List2.AddItem "<h1>"
List2.AddItem "<h2>"
List2.AddItem "<h3>"

End Sub

Private Sub Command10_Click()
frmBrowser.brwWebBrowser.Navigate App.Path & "\dropdown source.txt"
frmBrowser.WindowState = vbMaximized

End Sub


Private Sub Command11_Click()

End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command3_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\redirect.htm"
frmBrowser.WindowState = vbMaximized
List2.Clear
List2.AddItem "click the 8 to view the source!"
End Sub

Private Sub Command4_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\redirect source.txt"
frmBrowser.WindowState = vbMaximized
End Sub

Private Sub Command5_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\hover_a.htm"
frmBrowser.WindowState = vbMaximized
List2.Clear
List2.AddItem "click the 8 to view the source!"
End Sub

Private Sub Command6_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\hover_a source.txt"
frmBrowser.WindowState = vbMaximized

End Sub

Private Sub Command7_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\popupwin.htm"
frmBrowser.WindowState = vbMaximized
List2.Clear
List2.AddItem "click the 8 to view the source!"
End Sub

Private Sub Command8_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\popupwin source.txt"
frmBrowser.WindowState = vbMaximized

End Sub

Private Sub Command9_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\dropdown.htm"
frmBrowser.WindowState = vbMaximized
List2.Clear
List2.AddItem "click the 8 to view the source!"
End Sub

Private Sub list2_DblClick()
On Error Resume Next
Dim ggg As String
        ggg = "" & Form2.List2.Text & ""
         
    Mainer.ActiveForm.rtb1.SelText = "" & ggg & ""
      MsgBox " The Tag You Selected Was Inserted "
      Call HtmlHighlight
End Sub

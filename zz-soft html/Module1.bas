Attribute VB_Name = "public_declarations"
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

'**************
'*Declarations*
'**************
Public timei As String
Public Declare Function SendMessage Lib "User32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long


Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63
Public appversion As String



Public Const GWL_HWNDPARENT = (-8)


Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
       
'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public SelStart As Long  'start position in text box
Public TextLen As Long
Public Text As String

Public Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA

'these are just a bugs of variables for use
Public saved As String
Public pagenumber As Integer
Public insent As String
Public finish As String
Public typer  As String
Public sizer As String
Public U As String
Public i As String
Public b As String
Public colo As String
Public ue As String
Public ie As String
Public be As String
Public background As String
Public font As String
Public backgroundtext As String
Public user As String
Public fonter As String
Public Ruser As String
Public rcompany As String
Public retval As String

'this loads the users settings

Function loadset()
On Error Resume Next
 Open App.Path & "\settings.zdll" For Input As #1 'settings shold be in settings.zdll
       Line Input #1, background 'for form backcolor
       Line Input #1, backgroundtext 'for forms rtf backcolor
       Line Input #1, fonter 'this is only used in VWS , but has been diabled completely
       Line Input #1, user 'for user
     Close #1
     'this actually sets all the colors
     Form3.Show
          Form3.maper.BackColor = "" & backgroundtext & ""
          Form3.Hide
     Mainer.BackColor = "" & background & ""
frmBrowser.BackColor = "" & background & ""
Form4.BackColor = "" & background & ""
'the back color for form 2 keeps messing , to be fixed in v.2.1
Form2.BackColor = "" & background & ""
Form1.BackColor = "" & background & ""
'***
frmBrowser.Hide
Form4.Hide
Form2.Hide

End Function

Public Function RevInStr(ByVal sIn As String, sFind As String, Optional nStart As Long = 1, Optional bCompare As VbCompareMethod = vbBinaryCompare) As Long
Dim nPos As Long
    nPos = InStr(nStart, sIn, sFind, bCompare)
    If nPos = 0 Then
        RevInStr = 0
    Else
        RevInStr = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function

Public Function gobar()
Dim gobaroo As Integer
gobaroo = 0
Mainer.Pb.Value = 0
Mainer.Pb.Max = 200
Do Until gobaroo = 200
gobaroo = gobaroo + 1
Mainer.Pb.Value = Mainer.Pb.Value + 1
Loop
gobaroo = 0
Do Until gobaroo = 200
gobaroo = gobaroo + 1
Mainer.Pb.Value = Mainer.Pb.Value - 1
Loop
gobaroo = 0
Mainer.Pb.Value = 0
Mainer.Pb.Max = 200
End Function



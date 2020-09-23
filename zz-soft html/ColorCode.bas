Attribute VB_Name = "syntax"
' the code for syntax was used in part from another sytanxer , i forget its name


Option Explicit


Public m_TextCol As String
Public m_AttribCol As String
Public m_TagCol As String
Public m_CommentCol As String
Public m_AspCol As String


Public Sub HtmlHighlight()
On Error Resume Next
    HtmlColorCode
        Mainer.ActiveForm.rtb1.SelStart = 0
  End Sub


Public Function KeyPressEvent(KeyAscii As Integer) As Integer
    Static cInAttrib As Boolean, cInTag As Boolean
    Static cInAttribQuote As Boolean, cTypedIn As Boolean
    Static cInComment As Boolean
    Static cInASP As Boolean
    Static cInFunction As Boolean
    
    
    
    Dim cChar As String

    With Mainer.ActiveForm.rtb1
        cChar = Chr$(KeyAscii)
        
        If cInTag = False And cInAttrib = False And cInComment = False And cInASP = False Then
            .SelColor = m_TextCol
        End If

        If cInTag = True And (cInAttrib = True Or cInAttribQuote = True) Then
            .SelColor = m_AttribCol
        End If

        If cChar = "<" Then
            .SelColor = m_TagCol
            cInTag = True
            cTypedIn = True
        End If

        If cChar = "=" And cInTag = True Then
            cInAttrib = True
        End If

        If cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = True Then
            cInAttrib = False
            cInAttribQuote = False
        ElseIf cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = False Then
            cInAttribQuote = True
        End If

        If cChar = " " And (cInAttribQuote = False And cInTag = True) Then
            .SelColor = m_TagCol
            cInAttrib = False
        End If

        If cChar = "!" And Mid$(.Text, .SelStart, 1) = "<" Then

            .SelStart = .SelStart - 1
            .SelLength = 1
            .SelColor = m_CommentCol
            .SelText = "<!--"

            cInTag = False
            cInAttrib = False
            cInASP = False
            cInComment = True

            KeyAscii = 0
        End If
        
        If cChar = "%" And Mid$(.Text, .SelStart, 1) = "<" Then

            .SelStart = .SelStart - 1
            .SelLength = 1
            .SelColor = m_AspCol
            .SelText = "<%"

            cInTag = False
            cInAttrib = False
            cInASP = True
            cInComment = False

            KeyAscii = 0
        End If

        If cChar = ">" Then
            If cInComment = False And cInASP = True Then
                .SelColor = m_AspCol
            ElseIf cInComment = True And cInASP = False Then
                .SelColor = m_CommentCol
            ElseIf cInComment = False And cInASP = False Then
                .SelColor = m_TagCol
            End If
            
            cInTag = False
            cInASP = False
            cInComment = False
            cTypedIn = False
        End If

    End With

    KeyPressEvent = KeyAscii
    
   
ErrExit:
    Exit Function
End Function



Public Sub InsertTag(Tag$, StopAsp As Boolean)
Dim S As Long

    S = Mainer.ActiveForm.rtb1.SelStart
    If Len(Mainer.ActiveForm.rtb1.SelText) > 0 Then Mainer.ActiveForm.rtb1.SelText = ""
    Mainer.ActiveForm.rtb1.SelText = Tag$
    
    If StopAsp = True Then
 
        HtmlColorCode S, S + Len(Tag), True
        
    Else
     
        HtmlColorCode S, S + Len(Tag), False
     
    End If
    

End Sub



Public Sub InsertAspTag(Tag$)
Dim U As Long
    U = Mainer.ActiveForm.rtb1.SelStart
    If Len(Mainer.ActiveForm.rtb1.SelText) > 0 Then Mainer.ActiveForm.rtb1.SelText = ""
    Mainer.ActiveForm.rtb1.SelText = Tag$
    
    
    ASPColorCode U, U + Len(Tag)

End Sub


Public Function IsOutsideTag()
On Error Resume Next
Dim LastGT As Long, LastLT As Long, NextGT As Long, NextLT As Long
Dim EndTag As Long, StartTag As Long
Dim txt$, Start As Long, Start2 As Long
Dim InMainTag As Boolean, InEndTag As Boolean
    
    txt = Mainer.ActiveForm.rtb1.Text
    Start = Mainer.ActiveForm.rtb1.SelStart
    
    If Start = 0 Then
        m_TextCol = vbYellow
        Exit Function
    Else
        EndTag = InStr(Start + 1, txt, ">")
        StartTag = InStr(Start + 1, txt, "<")

        If StartTag > EndTag Then
            InMainTag = True
        Else
            InMainTag = False
        End If
        
        LastLT = RevInStr(txt, "<", Start + 1)
        LastGT = RevInStr(txt, ">", Start + 1)

        If LastLT < LastGT Then
            InEndTag = True
        Else
            InEndTag = False
        End If

        If InMainTag = True Or InEndTag = True Then
            m_TextCol = Mainer.ActiveForm.rtb1.SelColor
        Else
            m_TextCol = vbYellow
        End If
    End If
End Function

Public Function HtmlColorCode(Optional startchar As Long = 1, Optional endchar As Long = -1, Optional StopAsp As Boolean = False)
On Error GoTo ErrHandler
 
    Dim CommentOpenTag As String
    Dim CommentCloseTag As String

    Dim oldselstart As Long, oldsellen As Long
    

    Dim tag_open As Long
    Dim tag_close As Long
    Dim bef As String
    Dim Curr As String
    
   
    
    oldselstart = Mainer.ActiveForm.rtb1.SelStart
    oldsellen = Mainer.ActiveForm.rtb1.SelLength
    
    If endchar = -1 Then endchar = Len(Mainer.ActiveForm.rtb1.Text)
    If startchar = 0 Then startchar = 1

 
    
    tag_close = startchar
    
 
Mainer.ActiveForm.rtb1.HideSelection = True
    
  
    Do
      
        tag_open = InStr(tag_close, Mainer.ActiveForm.rtb1.Text, "<")
        
 
        If tag_open <> 0 Then
            
        
            bef = Mid$(Mainer.ActiveForm.rtb1.Text, 1, tag_open - 1)
            
      
            tag_close = InStr(tag_open, Mainer.ActiveForm.rtb1.Text, ">")

   
            Curr = Mid$(Mainer.ActiveForm.rtb1.Text, tag_open, tag_close - tag_open + 1)
            
            If tag_close <> 0 Then
                Select Case Left$(Curr, 3)
                    Case "<!-"
                 
                        tag_close = InStr(tag_open, Mainer.ActiveForm.rtb1.Text, "->") + 1
                            Mainer.ActiveForm.rtb1.SelStart = tag_open - 1
                           Mainer.ActiveForm.rtb1.SelLength = tag_close - tag_open + 1
                            Mainer.ActiveForm.rtb1.SelColor = m_CommentCol
                    Case Else
                 
                        cycleAttrib Curr, tag_open, tag_close
                End Select
            End If
            
            If tag_close = 0 Or tag_close >= endchar Then
    
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    

    If StopAsp = False Then
        ASPColorCode startchar, endchar
    End If
    
    Mainer.ActiveForm.rtb1.SelStart = oldselstart
Mainer.ActiveForm.rtb1.SelLength = oldsellen
  Mainer.ActiveForm.rtb1.HideSelection = False
  Mainer.ActiveForm.rtb1.SetFocus
    
  
    Exit Function
    
ErrHandler:
    Exit Function
End Function



Private Function ASPColorCode(Optional startchar As Long = 1, Optional endchar As Long = -1)
On Error GoTo ErrHandler
    Dim oldselstart As Long, oldsellen As Long
    
  
    Dim tag_open As Long
    Dim tag_close As Long
    Dim bef As String
    Dim Curr As String
    
   
    
    
    oldselstart = Mainer.ActiveForm.rtb1.SelStart
    oldsellen = Mainer.ActiveForm.rtb1.SelLength
    
    If endchar = -1 Then endchar = Len(Mainer.ActiveForm.rtb1.Text)
    If startchar = 0 Then startchar = 1


    
    tag_close = startchar
    
   
   Mainer.ActiveForm.rtb1.HideSelection = True
    
   
    Do
    
        tag_open = InStr(tag_close, Mainer.ActiveForm.rtb1.Text, "<%")
        
   
        If tag_open <> 0 Then
            
            
            bef = Mid$(Mainer.ActiveForm.rtb1.Text, 1, tag_open - 1)
            
         
            tag_close = InStr(tag_open, Mainer.ActiveForm.rtb1.Text, "%>")

           
            Curr = Mid$(Mainer.ActiveForm.rtb1.Text, tag_open, tag_close - tag_open + 1)
            
            If tag_close <> 0 Then
                Select Case Left$(Curr, 2)
                    Case "<%"
                   
                        tag_close = InStr(tag_open, Mainer.ActiveForm.rtb1.Text, "%>") + 1
                           Mainer.ActiveForm.rtb1.SelStart = tag_open - 1
                           Mainer.ActiveForm.rtb1.SelLength = tag_close - tag_open + 1
                           Mainer.ActiveForm.rtb1.SelColor = m_AspCol
                    Case Else
                      
                End Select
            End If
            
            If tag_close = 0 Or tag_close >= endchar Then
          
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
  Mainer.ActiveForm.rtb1.SelStart = oldselstart
 Mainer.ActiveForm.rtb1.SelLength = oldsellen
   Mainer.ActiveForm.rtb1.HideSelection = False
   Mainer.ActiveForm.rtb1.SetFocus
   
    
    Exit Function
    
ErrHandler:
    Exit Function
End Function

Private Function cycleAttrib(CurrTag As String, opentag As Long, closetag As Long)
    
    Dim fPos As Long, sPos As Long, qPos As Long, qnPos As Long, aPos As Long, tBeg As Long, tEnd As Long
    Dim isFirstCycle As Boolean
    Dim eTag As String
    Dim sPosTxt As String
    Dim LeftOver As Long
    Dim EndTag As Long, QuotePos As Long, QuoteEndPos As Long
    
    
    
    eTag = CurrTag
    isFirstCycle = True

    Do While Len(eTag) > 0
        fPos = InStr(1, eTag, "=")

        If (fPos = 0 And isFirstCycle = True) Then
        
            Mainer.ActiveForm.rtb1.SelStart = opentag - 1
            Mainer.ActiveForm.rtb1.SelLength = closetag - opentag + 1
            Mainer.ActiveForm.rtb1.SelColor = m_TagCol
            Exit Function
  
        ElseIf fPos <> 0 Then
            If Left$(eTag, 1) = "<" Then
             
                tBeg = opentag
                tEnd = opentag + fPos

             
                Mainer.ActiveForm.rtb1.SelStart = tBeg - 1
                Mainer.ActiveForm.rtb1.SelLength = closetag - tBeg + 1
               Mainer.ActiveForm.rtb1.SelColor = m_TagCol

             
                eTag = Mid$(eTag, fPos + 1)
                LeftOver = closetag - Len(eTag)
            End If
        End If
        
        
        sPos = InStr(1, eTag, Chr$(32))

     
        sPosTxt = Mid$(eTag, 1, sPos)
        
       
        qPos = InStr(1, sPosTxt, Chr$(34))

   
        If qPos <> 0 Then
            
            qnPos = InStr(2, eTag, Chr$(34))

            If qnPos <> 0 Then
                sPosTxt = Mid$(eTag, 1, qnPos)
            End If
        End If

        LeftOver = closetag - Len(eTag)
       Mainer.ActiveForm.rtb1.SelStart = LeftOver
       Mainer.ActiveForm.rtb1.SelLength = Len(sPosTxt)
      Mainer.ActiveForm.rtb1.SelColor = m_AttribCol
        
      
        eTag = Mid$(eTag, Len(sPosTxt) + 1)

       
        sPos = InStr(1, eTag, "=")

       
        If sPos = 0 Then
      
            eTag = Mid$(eTag, 1, Len(eTag) - 1)

            
      Mainer.ActiveForm.rtb1.SelStart = LeftOver
          Mainer.ActiveForm.rtb1.SelLength = Len(eTag)
          Mainer.ActiveForm.rtb1.SelColor = m_AttribCol

       
            sPos = Len(eTag)
            Exit Do
        End If

      
        eTag = Mid$(eTag, sPos + 1)
        isFirstCycle = False

     
        If sPos = 0 And qPos = 0 Then Exit Do
    Loop
    
  
    Exit Function
End Function


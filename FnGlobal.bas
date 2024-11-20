Attribute VB_Name = "FnGlobal"
Option Explicit

Public Function Replace(ByVal strSource As String, ByVal strWhat As String, ByVal strWith As String) As String

    Dim intstrWhatLength As Integer
    Dim intReplaceLength As Integer
    Dim intStart As Integer
    
    intstrWhatLength = Len(strWhat)
    
    If intstrWhatLength = 0 Then
        Replace = strSource
        Exit Function
    End If
    
    intReplaceLength = Len(strWith)
    intStart = InStr(1, strSource, strWhat)
    
    Do While intStart > 0
        strSource = Left(strSource, intStart - 1) + strWith + Right(strSource, Len(strSource) - (intStart + intstrWhatLength - 1))
        intStart = InStr(intStart + intReplaceLength, strSource, strWhat)
    Loop
    
    Replace = strSource

End Function

Function ToUnicode(UC As Long) As String
    ToUnicode = "\u" _
        & UnicodeNibble((UC / 4096) And &HF) _
        & UnicodeNibble((UC / 256) And &HF) _
        & UnicodeNibble((UC / 16) And &HF) _
        & UnicodeNibble(UC And &HF)
End Function

Function UnicodeNibble(Nibble As Long) As String
    UnicodeNibble = Mid("0123456789abcdef", 1 + Nibble, 1)
End Function

Public Function ToEscapedString(Value As String) As String

    Dim Temp As String
    Dim Indx As Integer
    Dim Chdx As Long
    
    For Indx = 1 To Len(Value)
    
        Chdx = AscW(Mid(Value, Indx, 1))
        
        If Chdx >= 0 And Chdx <= 127 Then
        
            Select Case Chdx
            Case 34 '\"
                Temp = Temp & "\"""
            Case 92 ' '\\
                Temp = Temp & "\\"
            Case 47 ' '\
                Temp = Temp & "\/"
            Case &H8 ' '\b
                Temp = Temp & "\b"
            Case &H9 ' '\t
                Temp = Temp & "\t"
            Case &HA ' '\n
                Temp = Temp & "\n"
            Case &HC ' '\f
                Temp = Temp & "\f"
            Case &HD ' '\r
                Temp = Temp & "\r"
            Case Else
                Temp = Temp & ChrW(Chdx)
            End Select
            
        Else
        
            Temp = Temp & ToUnicode(Chdx)
            
        End If
        
    Next Indx
        
    ToEscapedString = Temp
  
End Function


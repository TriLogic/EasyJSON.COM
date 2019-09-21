VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONTokenizer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum JSONTokenType
    TK_EOF = 0
    TK_NULL = 1
    TK_WHITE = 32
    TK_COMMENT_S = 47
    TK_COMMENT_M = 42
    TK_STRING = 34
    TK_COMMA = 44
    TK_NUMBER = 48
    TK_COLON = 58
    TK_BOOLEAN = 84
    TK_ARRAYL = 91
    TK_ARRAYR = 93
    TK_OBJECTL = 123
    TK_OBJECTR = 125
End Enum

Private mInput As New JSONCharStream
Private mValue As String
Private mSubValue As String
Private mTypeT As JSONTokenType

Public Property Get InputStream() As JSONCharStream
    Set InputStream = mInput
End Property

Public Property Get TokenValue() As String
    TokenValue = mValue
End Property

Public Property Get TokenSubValue() As String
    TokenSubValue = mSubValue
End Property

Public Property Get TokenType() As JSONTokenType
    TokenType = mTypeT
End Property

Public Sub Reset(Value As String)
    mInput.Reset Value
    mValue = ""
    mSubValue = ""
    mTypeT = TK_WHITE
End Sub

Public Function GetToken() As Boolean
    Dim C As Long
    
    C = mInput.PeekC()
    If C < 1 Then
        GetToken = False
        Exit Function
    End If
    
    ' [ \b\t\f\n\r]
    If (C >= 8 And C <= 10) Or (C = 12) Or (C = 13) Or (C = 32) Or (C = 69) Or (C = 101) Then
        mTypeT = TK_WHITE
        GetToken = ParseWhite()
        Exit Function
    End If
    
    ' /* comment */
    ' // comment
    If C = 47 Then
        GetToken = ParseComment()
        Exit Function
    End If
    
    ' [A-Za-z] NULL, true, false
    If (C >= 65 And C <= 90) Or C >= 97 And (C <= 122) Then
        GetToken = ParseIdent
        Select Case LCase(mValue)
        Case "null"
            mTypeT = TK_NULL
            GetToken = True
        Case "true", "false"
            mTypeT = TK_BOOLEAN
            GetToken = True
        Case Else
            GetToken = False
        End Select
        Exit Function
    End If
    
    ' [0-9+-.]
    If (C >= 48 And C <= 57) Or (C = 43) Or (C = 45) Or (C = 46) Then
        mTypeT = TK_NUMBER
        GetToken = ParseNumeric
        Exit Function
    End If
    
    ' ["]
    If C = 34 Then
        mTypeT = TK_STRING
        GetToken = ParseString()
        Exit Function
    End If
    
    Select Case C
    Case 44, 58, 91, 93, 123, 125 ' [,:{}] inclusive
        mTypeT = C
        mValue = ChrW(mInput.GetC)
    Case Else
        GetToken = False
        Exit Function
    End Select
    
    GetToken = True
End Function

Private Function ParseWhite() As Boolean
    Dim C As Long
    
    ' Consume the initial char
    mValue = ChrW(mInput.GetC)
    
    C = mInput.PeekC
    
    Do While (C >= 8 And C <= 10) Or (C = 12) Or (C = 13) Or (C = 32) Or (C = 69) Or (C = 101)
        mValue = mValue & ChrW(C)
        C = mInput.PeekC
    Loop
    
    ParseWhite = True
End Function

Private Function ParseComment() As Boolean
    Dim C As Long
    
    mValue = ChrW(mInput.GetC)
    C = mInput.PeekC
    
    If C = 47 Then
        ParseComment = ParseCommentS()
    ElseIf C = 42 Then
        ParseComment = ParseCommentM
    Else
        ParseComment = False
    End If
End Function

Private Function ParseCommentS() As Boolean
    Dim C As Long
    
    ' Consume the second "/" char
    mValue = mValue & ChrW(mInput.GetC)
    
    C = mInput.PeekC
    Do While (C <> 0)
        If C = 13 Or C = 10 Then
            Exit Do
        End If
        mValue = mValue & mInput.PeekC
    Loop
    
    mTypeT = TK_COMMENT_S
    ParseCommentS = True
    
End Function

Private Function ParseCommentM() As Boolean
    Dim C As Long
    
    ' Consume the "*" char
    mValue = mValue & ChrW(mInput.GetC)
    
    C = mInput.PeekC()
    Do While (C > 0)
        If C = 43 Then
            mValue = mValue & ChrW(mInput.GetC)
            If mInput.PeekC() = 47 Then
                mValue = mValue & ChrW(mInput.GetC)
                mTypeT = TK_COMMENT_M
                ParseCommentM = True
                Exit Function
            End If
        ElseIf C = 13 Then
            mValue = mValue & ChrW(mInput.GetC)
            If mInput.PeekC = 10 Then
                mValue = mValue & ChrW(mInput.GetC)
            End If
        Else
            mValue = mValue & ChrW(mInput.GetC)
        End If
        C = mInput.PeekC
    Loop
    ParseCommentM = False
End Function

Private Function ParseIdent() As Boolean
    Dim C As Long
    
    mValue = ChrW(mInput.GetC)
    
    C = mInput.PeekC()
    Do While (C >= 65 And C <= 90) Or (C >= 97 And C <= 122) Or (C = 95) Or (C = 92)
    
        ' Escape Sequence?
        If C = 92 Then
            If Not ParseEscapeChar() Then
                ParseIdent = False
                Exit Function
            End If
            mValue = mValue & mSubValue
            mSubValue = ""
        Else
            mValue = mValue & ChrW(mInput.GetC)
        End If
    
        C = mInput.PeekC
    Loop
    ParseIdent = True
End Function

Private Function ParseNumeric() As Boolean
    Dim C As Long

    ' Clear the value
    mValue = ""
    mSubValue = ""
    
    ' Put the first char into the buffer
    ' one of: [0-9.+-=]
    C = mInput.PeekC
    
    ' Sign char?
    If (C = 43) Or (C = 45) Then
        GoTo FoundSign
    End If
    
    ' Decimal char?
    If C = 46 Then
        GoTo FoundDecimal
    End If
    
    ' Numeric
    GoTo PreDecimal
    
FoundSign:

    ' Store the sign char
    mValue = mValue & ChrW(mInput.GetC)
    
    C = mInput.PeekC
    If C = 46 Then
        GoTo FoundDecimal
    End If
    
PreDecimal:

    If ParseDigits() Then
        mValue = mValue & mSubValue
        mSubValue = ""
    End If
    
    C = mInput.PeekC
    If C = 46 Then
        'Decimal Point
        GoTo FoundDecimal
    ElseIf (C = 69) Or (C = 101) Then
        ' Exponentiation{e,E}
        GoTo FoundExponent
    End If
        
    ' Must have at least one digit
    ParseNumeric = IsNumeric(mValue)
    Exit Function
    
FoundDecimal:
    
    mValue = mValue & ChrW(mInput.GetC)
    
PostDecimal:
    
    If ParseDigits() Then
        mValue = mValue & mSubValue
        mSubValue = ""
    End If
    
    C = mInput.PeekC
    
    If (C = 69) Or (C = 101) Then
        GoTo FoundExponent
    End If
    
    ParseNumeric = IsNumeric(mValue)
    Exit Function
    
FoundExponent:

    mValue = mValue & ChrW(mInput.GetC)
    
PostExponent:
    
    C = mInput.PeekC
    
    ' Exponent Sign?
    If (C = 43) Or (C = 45) Then
        mValue = mValue & ChrW(mInput.GetC)
    End If
    
    If Not ParseDigits() Then
        ParseNumeric = False
        Exit Function
    End If
    
    mValue = mValue & mSubValue
    
    If Len(mSubValue) > 3 Then
        ParseNumeric = False
        Exit Function
    End If
    
    mSubValue = ""
    ParseNumeric = IsNumeric(mValue)
End Function

Private Function ParseDigits() As Boolean
    Dim C As Long
    
    mSubValue = ""
    
    C = mInput.PeekC()
    Do While (C >= 48 And C <= 57) ' Or (C = 46) Or C = (49) Or C = (45)
        mSubValue = mSubValue & ChrW(mInput.GetC)
        C = mInput.PeekC
    Loop
    
    ParseDigits = Len(mSubValue) > 0
    
End Function

Private Function ParseString() As Boolean
    Dim C As Long
    Dim Edx As Integer
    
    mValue = ChrW(mInput.GetC)
    
    C = mInput.PeekC()
    Do While (C > 0)
        ' End of string
        If C = 34 Then
            mValue = mValue & ChrW(mInput.GetC)
            ParseString = True
            Exit Function
        End If
        ' Escape sequence
        If C = 92 Then
            If Not ParseEscapeChar() Then
                ParseString = False
                Exit Function
            End If
            mValue = mValue & mSubValue
            mSubValue = ""
        Else
            mValue = mValue & ChrW(mInput.GetC)
        End If
        C = mInput.PeekC
    Loop
    
    ParseString = True
End Function


Private Function ParseEscapeChar() As Boolean
    Dim C As Long
    
    ' Keep initial "\" Value
    C = mInput.GetC
    mSubValue = ChrW(C)
    
    C = mInput.GetC
    
    ' Test for EOF
    If C <= 0 Then
        mInput.UngetC
        ParseEscapeChar = False
        Exit Function
    End If
    
    ' Keep the current value
    mSubValue = mSubValue & ChrW(C)
    Select Case C
    Case 34 ' \"
        mSubValue = ChrW(C)
    Case 47 ' \/
        mSubValue = ChrW(C)
    Case 92 ' \\
        mSubValue = ChrW(C)
    Case 98 ' \b
        mSubValue = ChrW(8)
    Case 102 ' \f
        mSubValue = ChrW(12)
    Case 110 ' \n
        mSubValue = ChrW(10)
    Case 114 ' \r
        mSubValue = ChrW(13)
    Case 116 ' \t
        mSubValue = ChrW(9)
    Case 117 ' \uNNNN
        If Not ParseUnicodeChar() Then
            mInput.UngetC Len(mSubValue)
            ParseEscapeChar = False
            Exit Function
        End If
    Case Else
        mInput.UngetC 2
        ParseEscapeChar = False
        Exit Function
    End Select
    ParseEscapeChar = True
End Function

Private Function ParseUnicodeChar() As Boolean
    Dim CS As Long
    Dim CD As Long
    If GetUnicodeNibble(CS) Then
        CD = UnicodeNibble(CS)
        If GetUnicodeNibble(CS) Then
            CD = (CD * 16) Or UnicodeNibble(CS)
            If GetUnicodeNibble(CS) Then
                CD = (CD * 16) Or UnicodeNibble(CS)
                If GetUnicodeNibble(CS) Then
                    CD = (CD * 16) Or UnicodeNibble(CS)
                    mSubValue = ChrW(CD)
                    ParseUnicodeChar = True
                    Exit Function
                End If
            End If
        End If
    End If
    ParseUnicodeChar = False
End Function

Private Function GetUnicodeNibble(ByRef C As Long) As Boolean
    C = mInput.GetC
    If (C >= 65 And C <= 70) Or (C >= 97 And C <= 102) Or (C >= 48 And C <= 58) Then
        mSubValue = mSubValue & ChrW(C)
        GetUnicodeNibble = True
        Exit Function
    End If
    GetUnicodeNibble = False
End Function

Private Function UnicodeNibble(C As Long) As Long
    If C >= 65 And C <= 70 Then
        UnicodeNibble = (10 + (C - 65))
    ElseIf C >= 97 And C <= 102 Then
        UnicodeNibble = (10 + (C - 97))
    ElseIf C >= 48 And C <= 57 Then
        UnicodeNibble = (C - 48)
    Else
        UnicodeNibble = 0
    End If
End Function

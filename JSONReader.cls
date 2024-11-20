VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private mItem As JSONItem
Private mTokens As JSONTokenizer

Public Property Get Tokenizer() As JSONTokenizer
    Set Tokenizer = mTokens
End Property

Public Property Get Item() As JSONItem
    Set Item = mItem
End Property

Public Function GetObject(json As String) As JSONItem

    Set mItem = Nothing
    Set mTokens = New JSONTokenizer
    mTokens.Reset json

    If ExpectToken(TK_OBJECTL) Then
        
        Call ParseObject
        
        If mItem.Parent Is Nothing Then
            Set GetObject = mItem
            Exit Function
        End If
        
        Err.Raise 66001, , "JSON: Invalid Object"
        
    End If
        
    Err.Raise 66001, , "JSON: Expected Object"

End Function

Private Sub ParseObject(Optional Key As String = "")

    Dim More As Boolean
    Dim oItem As JSONObject
    
    Set oItem = New JSONObject
    If mItem Is Nothing Then
        Set mItem = oItem
    Else
        mItem.Add oItem, Key
        Set mItem = oItem
    End If
    
    More = ParseObjectItem()
    Do While More
    
        If Not ExpectToken2(TK_COMMA, TK_OBJECTR) Then
            Err.Raise 66001, , "JSON: Invalid Object Syntax"
        End If
        
        If mTokens.TokenType = TK_OBJECTR Then
            Exit Do
        End If
    
        More = ParseObjectItem()
    Loop

    If mTokens.TokenType <> TK_OBJECTR Then
        Err.Raise 66001, , "JSON: Invalid Object Syntax"
    End If

    If Not (mItem.Parent Is Nothing) Then
        Set mItem = mItem.Parent
    End If

End Sub

Private Sub ParseArray(Optional Key As String = "")

    Dim More As Boolean
    Dim oItem As JSONItem
    
    Set oItem = New JSONArray
    mItem.Add oItem, Key
    Set mItem = oItem
    
    More = ParseSubItem(TK_ARRAYR)
    Do While More
    
        If Not ExpectToken2(TK_COMMA, TK_ARRAYR) Then
            Err.Raise 66001, , "JSON: Array Object Syntax"
        End If
        
        If mTokens.TokenType = TK_ARRAYR Then
            Exit Do
        End If
    
        More = ParseSubItem(TK_ARRAYR)
    Loop
    
    If mTokens.TokenType <> TK_ARRAYR Then
        Err.Raise 66001, , "JSON: Invalid Array Syntax"
    End If

    Set mItem = mItem.Parent

End Sub

Private Function ParseSubItem(AcceptableExit As JSONTokenType, Optional Key As String = "") As Boolean

    If Not GetToken() Then
        If AcceptableExit = TK_ARRAYR Then
            Err.Raise 66001, , "JSON: Invalid Array Syntax"
        Else
            Err.Raise 66001, , "JSON: Invalid Object Syntax"
        End If
    End If
    
    If mTokens.TokenType = AcceptableExit Then
        ParseSubItem = False
        Exit Function
    End If
    
    Select Case mTokens.TokenType
    Case TK_BOOLEAN
        mItem.AddBoolean LCase(mTokens.TokenValue) = "true", Key
    Case TK_NUMBER
        mItem.AddNumber CDbl(mTokens.TokenValue), Key
    Case TK_STRING
        mItem.AddString StripQuotes(mTokens.TokenValue), Key
    Case TK_NULL
        mItem.AddNull Key
    Case TK_ARRAYL
        ParseArray Key
    Case TK_OBJECTL
        ParseObject Key
    Case Else
        Err.Raise 66001, , "JSON: Invalid Array Syntax"
    End Select
        
    ParseSubItem = True

End Function

Private Function ParseObjectItem() As Boolean
    Dim sName As String
    Dim sTemp As String

    If Not ExpectToken2(TK_STRING, TK_OBJECTR) Then
        Err.Raise 66001, , "JSON: Invalid Object Syntax"
    End If
    
    If mTokens.TokenType = TK_OBJECTR Then
        ParseObjectItem = False
        Exit Function
    End If
    
    sName = StripQuotes(mTokens.TokenValue)
    
    If mItem.Exists(sName) Then
        Err.Raise 66001, , "JSON: Duplicate Object Key"
    End If
    
    If Not ExpectToken(TK_COLON) Then
        Err.Raise 66001, , "JSON: Invalid Object Syntax"
    End If

    If Not ParseSubItem(TK_OBJECTR, sName) Then
        Err.Raise 66001, , "JSON: Invalid Object Syntax"
    End If
    
    ParseObjectItem = True

End Function

Private Function GetToken() As Boolean
    Do While mTokens.GetToken()
        ' Skip whitespace and comments
        Select Case mTokens.TokenType
        Case TK_WHITE, TK_COMMENT_S, TK_COMMENT_M
            ' do nothing
        Case Else
            GetToken = True
            Exit Function
        End Select
    Loop
    GetToken = False
End Function

Private Function ExpectToken(TokenType As JSONTokenType) As Boolean
    If GetToken() Then
        ExpectToken = mTokens.TokenType = TokenType
        Exit Function
    End If
    ExpectToken = False
End Function

Private Function ExpectToken2(TokenType1 As JSONTokenType, TokenType2 As JSONTokenType) As Boolean
    If GetToken() Then
        ExpectToken2 = (mTokens.TokenType = TokenType1) Or (mTokens.TokenType = TokenType2)
        Exit Function
    End If
    ExpectToken2 = False
End Function

Private Function ExpectToken3(TokenType1 As JSONTokenType, TokenType2 As JSONTokenType, TokenType3 As JSONTokenType) As Boolean
    If GetToken() Then
        ExpectToken3 = (mTokens.TokenType = TokenType1) Or (mTokens.TokenType = TokenType2) Or (mTokens.TokenType = TokenType3)
        Exit Function
    End If
    ExpectToken3 = False
End Function

Private Function StripQuotes(Value As String) As String
    If Value = "" Or Value = """" Then
        StripQuotes = ""
    Else
        Value = Left(Value, Len(Value) - 1)
        Value = Right(Value, Len(Value) - 1)
    End If
    StripQuotes = Value
End Function


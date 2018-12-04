Attribute VB_Name = "Module1"
Option Explicit

Sub Main()

    Dim Result As String

    'Call TestBuilder(True)
    ' Call TestTokenizer(True)
    ' Call TestReader(True)
    Call TestWriter(True)

End Sub

Public Function TestReader(Optional DPrint As Boolean = False) As Boolean

    Dim oRead As New JSONReader
    Dim oItem As JSONItem
    Dim iJson As String
    Dim oJson As String
    
    iJson = TestBuilder(DPrint)
    Set oItem = oRead.GetObject(iJson)
    
    oJson = oItem.ToString()
    If DPrint Then
        Debug.Print "JSON:" & oJson
    End If
    

End Function

Public Function TestTokenizer(Optional DPrint As Boolean = False) As Boolean

    Dim tokens As New JSONTokenizer
    Dim json As String
    json = TestBuilder(DPrint)
    
    tokens.Reset json
    
    Do While tokens.GetToken()
        Debug.Print "TOKEN=""" & tokens.TokenValue & """"
    Loop

End Function

Function TestBuilder(Optional DPrint As Boolean = False) As String

    Dim json As String
    Dim O As JSONItem
    Set O = New JSONObject
    
    ' GoTo Numbers
    ' GoTo Empties

Strings:

    Set O = O.AddObject("Strings")
    O.AddString "", "empty"
    O.AddString vbCr, "vbCr"
    O.AddString vbLf, "vbLf"
    O.AddString vbTab, "vbTab"
    O.AddString Chr(8), "bsckSpace"
    O.AddString "\", "backSlash"
    O.AddString "/", "slash"
    O.AddString """", "dquote"
    O.AddString ChrW(&H110), "unicode"
    O.AddString "The quick silver fox jumped over the lazy brown dog.", "allChars"
    Set O = O.Parent

Constants:

    Set O = O.AddObject("Constants")
    O.AddBoolean True, "true"
    O.AddBoolean False, "false"
    O.AddNull "null"
    Set O = O.Parent
    
Numbers:
    
    Set O = O.AddObject("Numbers")
    O.AddNumber 0, "zero"
    O.AddNumber 1, "one"
    O.AddNumber -3, "negative"
    O.AddNumber 0.005, "decimal"
    O.AddNumber -5.79E-32, "scientific"
    Set O = O.Parent

Empties:

    Call O.AddObject("Empties") _
        .AddObject("HostObj") _
        .AddArray("Array") _
        .AddObject() _
        .AddNumber(1, "One") _
        .AddNumber(2, "Two") _
        .Parent _
        .AddObject() _
        .AddNumber(3, "Three") _
        .AddNumber(4, "Four") _
        .Parent _
        .Parent _
        .AddObject("Object") _
        .Parent _
        .Parent _
        .AddArray("HostArr") _
        .AddObject _
        .Parent _
        .AddArray

Done:

    json = O.ToString()
    If DPrint Then
        Debug.Print "JSON:" & json
    End If
    TestBuilder = O.ToString()

End Function

Public Function TestWriter(Optional DPrint As Boolean = True)

    Dim oWrite As JSONWriter
    Dim oRead As JSONReader
    Dim oItem As JSONItem
    
    Dim json As String
    json = TestBuilder(True)
    
    Set oRead = New JSONReader
    Set oItem = oRead.GetObject(json)

    Set oWrite = New JSONWriter
    oWrite.SetFormatAllman , True, True
    'oWrite.SetFormatKNR OutdentClose:=True
    'oWrite.ArrayArrayPfx = "%crlf%%>1%"
    'oWrite.ArrayObjectPfx = "%crlf%%>1%"
    'oWrite.SetFormatWhitesmith
    'oWrite.SetFormatLinear False
    'oWrite.SetFormatLinear True
    
    Debug.Print oWrite.ToString(oItem)

End Function



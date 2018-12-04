VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONString"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements JSONItem

Private mData As Variant
Private mParent As JSONItem

Public Sub Class_Initialize()
    mData = ""
End Sub

Public Property Get JSONItem_Count() As Integer
    JSONItem_Count = 0
End Property

Public Property Get JSONItem_Item(Key As String) As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Property

Public Property Get JSONItem_Exists(Key As String) As Boolean
    Err.Raise 66001, , "JSON: Invalid Container"
End Property

Public Function JSONItem_Add(Value As JSONItem, Optional Key As String = "") As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Function

Public Function JSONItem_AddArray(Optional Key As String = "") As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Function

Public Function JSONItem_AddObject(Optional Key As String = "") As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Function

Public Function JSONItem_AddNumber(Value As Double, Optional Key As String = "") As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Function

Public Function JSONItem_AddString(Value As String, Optional Key As String = "") As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Function

Public Function JSONItem_AddBoolean(Value As Boolean, Optional Key As String = "") As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Function

Public Function JSONItem_AddNull(Optional Key As String = "") As JSONItem
    Err.Raise 66001, , "JSON: Invalid Container"
End Function

Public Sub JSONItem_Remove(Key As String)
    Err.Raise 66001, , "JSON: Invalid Container"
End Sub

Public Property Get JSONItem_JsType() As String
    JSONItem_JsType = "string"
End Property

Public Property Get JSONItem_IsJsArray() As Boolean
    JSONItem_IsJsArray = False
End Property

Public Property Get JSONItem_IsJsBool() As Boolean
    JSONItem_IsJsBool = False
End Property

Public Property Get JSONItem_IsJsNull() As Boolean
    JSONItem_IsJsNull = False
End Property

Public Property Get JSONItem_IsJsNumber() As Boolean
    JSONItem_IsJsNumber = False
End Property

Public Property Get JSONItem_IsJsObject() As Boolean
    JSONItem_IsJsObject = False
End Property

Public Property Get JSONItem_IsJsString() As Boolean
    JSONItem_IsJsString = True
End Property

Public Property Get JSONItem_ToString() As String
    Dim Temp As String
    Dim Indx As Integer
    Dim Chdx As Long
    
    For Indx = 1 To Len(mData)
    
        Chdx = AscW(Mid(mData, Indx, 1))
        
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
    
    JSONItem_ToString = """" & Temp & """"
    
End Property

Private Function ToUnicode(UC As Long) As String
    ToUnicode = "\u" _
        & UnicodeNibble((UC / 4096) And &HF) _
        & UnicodeNibble((UC / 256) And &HF) _
        & UnicodeNibble((UC / 16) And &HF) _
        & UnicodeNibble(UC And &HF)
End Function

Private Function UnicodeNibble(Nibble As Long) As String
    UnicodeNibble = Mid("0123456789abcdef", 1 + Nibble, 1)
End Function

Public Sub JSONItem_SetValue(Value As Variant)
    mData = CStr(Value)
End Sub
Public Function JSONItem_GetValue() As Variant
    JSONItem_GetValue = CStr(mData)
End Function

Public Property Get JSONItem_Parent() As JSONItem
    Set JSONItem_Parent = mParent
End Property

Public Property Set JSONItem_Parent(Value As JSONItem)
    Set mParent = Value
End Property




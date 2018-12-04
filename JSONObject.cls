VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONObject"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements JSONItem

Private mData As Variant
Private mParent As JSONItem

Public Sub Class_Initialize()
    Set mData = New Dictionary
End Sub

Public Property Get JSONItem_Count() As Integer
    JSONItem_Count = mData.Count
End Property

Public Property Get JSONItem_Item(Key As String) As JSONItem
    If Key = "" Then
        Err.Raise 66001, , "JSON: Invalid Object Key"
    End If
    Set JSONItem_Item = mData(Key)
End Property

Public Property Get JSONItem_Exists(Key As String) As Boolean
    If Key = "" Then
        Err.Raise 66001, , "JSON: Invalid Object Key"
    End If
    JSONItem_Exists = mData.Exists(Key)
End Property

Public Function JSONItem_Add(Value As JSONItem, Optional Key As String = "") As JSONItem
    If (IsMissing(Key) Or Key = "") Then
        Err.Raise 66001, , "Invalid Object Key"
    End If
    If mData.Exists(Key) Then
        Err.Raise 66001, , "Duplicate Object Key"
    End If
    Set mData(Key) = Value
    Set Value.Parent = Me
    Set JSONItem_Add = IIf(Value.IsJsObject Or Value.IsJsArray, Value, Me)
End Function

Public Function JSONItem_AddArray(Optional Key As String = "") As JSONItem
    Set JSONItem_AddArray = JSONItem_Add(New JSONArray, Key)
End Function

Public Function JSONItem_AddObject(Optional Key As String = "") As JSONItem
    Set JSONItem_AddObject = JSONItem_Add(New JSONObject, Key)
End Function

Public Function JSONItem_AddNumber(Value As Double, Optional Key As String = "") As JSONItem
    Dim Item As JSONItem
    Set Item = New JSONNumber
    Item.SetValue Value
    Set JSONItem_AddNumber = JSONItem_Add(Item, Key)
End Function

Public Function JSONItem_AddString(Value As String, Optional Key As String = "") As JSONItem
    Dim Item As JSONItem
    Set Item = New JSONString
    Item.SetValue Value
    Set JSONItem_AddString = JSONItem_Add(Item, Key)
End Function

Public Function JSONItem_AddBoolean(Value As Boolean, Optional Key As String = "") As JSONItem
    Dim Item As JSONItem
    Set Item = New JSONBool
    Item.SetValue Value
    Set JSONItem_AddBoolean = JSONItem_Add(Item, Key)
End Function

Public Function JSONItem_AddNull(Optional Key As String = "") As JSONItem
    Set JSONItem_AddNull = JSONItem_Add(New JSONNull, Key)
End Function

Public Sub JSONItem_Remove(Key As String)
    Err.Raise 66001, , "JSON: Invalid Container"
End Sub

Public Property Get JSONItem_JsType() As String
    JSONItem_JsType = "object"
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
    JSONItem_IsJsObject = True
End Property

Public Property Get JSONItem_IsJsString() As Boolean
    JSONItem_IsJsString = False
End Property

Public Property Get JSONItem_ToString() As String
    Dim Result As String
    Dim Key As Variant
    Dim Item As JSONItem
    Dim OutCount As Integer
    Result = "{"
    For Each Key In mData.Keys
        Set Item = mData(Key)
        OutCount = OutCount + 1
        If OutCount > 1 Then
            Result = Result & ","
        End If
        Result = Result & """" & Key & """:" & Item.ToString()
    Next
    Result = Result & "}"
    JSONItem_ToString = Result
End Property

Public Sub JSONItem_SetValue(Value As Variant)
    Err.Raise 66001, , "JSON: Invalid Operation on Object"
End Sub
Public Function JSONItem_GetValue() As Variant
    Set JSONItem_GetValue = mData
End Function

Public Property Get JSONItem_Parent() As JSONItem
    Set JSONItem_Parent = mParent
End Property

Public Property Set JSONItem_Parent(Value As JSONItem)
    Set mParent = Value
End Property



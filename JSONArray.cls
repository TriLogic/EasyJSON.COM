VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements JSONItem

Private mData As Variant
Private mParent As JSONItem

Public Sub Class_Initialize()
    ReDim mData(0)
End Sub

Public Property Get JSONItem_Count() As Integer
    JSONItem_Count = UBound(mData)
End Property

Public Property Get JSONItem_Item(Key As String) As JSONItem
    If Key = "" Or (Not IsNumeric(Key)) Then
        Err.Raise 66001, , "JSON: Invalid Array Key"
    End If
    Set JSONItem_Item = mData(CInt(Key))
End Property

Public Property Get JSONItem_Exists(Key As String) As Boolean
    If Key = "" Or (Not IsNumeric(Key)) Then
        Err.Raise 66001, , "JSON: Invalid Array Key"
    End If
    Dim Idx As Long
    Idx = CLng(Key)
    JSONItem_Exists = Idx >= 1 And Idx <= UBound(mData)
End Property

Public Function JSONItem_Add(Value As JSONItem, Optional Key As String = "") As JSONItem
    ' if a key was supplied then we have an error
    If Key <> "" Then
        Err.Raise 66001, , "Invalid Operation Array Key"
    End If
    
    ' Add the item to the end
    ReDim Preserve mData(UBound(mData) + 1)
    Set mData(UBound(mData)) = Value
    Set Value.Parent = Me
    
    ' If the item is an array or object then return it, otherwise return myself
    Set JSONItem_Add = IIf(Value.IsJsObject Or Value.IsJsArray, Value, Me)
End Function

Public Function JSONItem_AddArray(Optional Key As String = "") As JSONItem
    If Key <> "" Then
        Err.Raise 66001, , "Invalid Operation Array Key"
    End If
    Set JSONItem_AddArray = JSONItem_Add(New JSONArray, Key)
End Function

Public Function JSONItem_AddObject(Optional Key As String = "") As JSONItem
    If Key <> "" Then
        Err.Raise 66001, , "Invalid Operation Array Key"
    End If
    
    Set JSONItem_AddObject = JSONItem_Add(New JSONObject, Key)
End Function

Public Function JSONItem_AddNumber(Value As Double, Optional Key As String = "") As JSONItem
    If Key <> "" Then
        Err.Raise 66001, , "Invalid Operation Array Key"
    End If
    
    Dim Item As JSONItem
    Set Item = New JSONNumber
    
    Item.SetValue Value
    
    Set JSONItem_AddNumber = JSONItem_Add(Item, Key)
End Function

Public Function JSONItem_AddString(Value As String, Optional Key As String = "") As JSONItem
    If Key <> "" Then
        Err.Raise 66001, , "Invalid Operation Array Key"
    End If
    
    Dim Item As JSONItem
    Set Item = New JSONString
    Item.SetValue Value
    
    Set JSONItem_AddString = JSONItem_Add(Item, Key)
End Function

Public Function JSONItem_AddBoolean(Value As Boolean, Optional Key As String = "") As JSONItem
    If Key <> "" Then
        Err.Raise 66001, , "Invalid Operation Array Key"
    End If
    
    Dim Item As JSONItem
    Set Item = New JSONBool
    Item.SetValue Value
    
    Set JSONItem_AddBoolean = JSONItem_Add(Item, Key)
End Function

Public Function JSONItem_AddNull(Optional Key As String = "") As JSONItem
    If Key <> "" Then
        Err.Raise 66001, , "Invalid Operation Array Key"
    End If
    
    Set JSONItem_AddNull = JSONItem_Add(New JSONNull, Key)
End Function

Public Sub JSONItem_Remove(Key As String)
    If Not JSONItem_Exists(Key) Then
        Err.Raise 66001, , "JSON: Invalid Array Key"
    End If
    
    Dim Idx As Long
    Dim Jdx As Long
    Dim mNew() As Variant
    ReDim mNew(0)
    
    Jdx = CLng(Key) + 1
    For Idx = 1 To UBound(mData)
        If Idx <> Jdx Then
            ReDim Preserve mNew(UBound(mNew) + 1)
            Set mNew(UBound(mNew)) = mData(Idx)
        End If
    Next Idx
    
    ReDim mData(0)
    
    mData = mNew
End Sub

Public Property Get JSONItem_ItemType() As JSONType
    JSONItem_ItemType = JSON_Array
End Property

Public Property Get JSONItem_IsJsArray() As Boolean
    JSONItem_IsJsArray = True
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
    JSONItem_IsJsString = False
End Property

Public Property Get JSONItem_ToString() As String
    Dim Key As Long
    Dim Result As String
    Dim Item As JSONItem
    
    Result = "["
    If UBound(mData) > 0 Then
        Set Item = mData(1)
        Result = Result & Item.ToString()
        For Key = 2 To UBound(mData)
        Set Item = mData(Key)
            Result = Result & "," & Item.ToString()
        Next Key
    End If
    
    Result = Result & "]"
    JSONItem_ToString = Result
End Property

Public Sub JSONItem_SetValue(Value As Variant)
    Err.Raise 66001, , "JSON: Invalid Operation on Array"
End Sub

Public Function JSONItem_GetValue() As Variant
    JSONItem_GetValue = mData
End Function

Public Property Get JSONItem_Parent() As JSONItem
    Set JSONItem_Parent = mParent
End Property

Public Property Set JSONItem_Parent(Value As JSONItem)
    Set mParent = Value
End Property




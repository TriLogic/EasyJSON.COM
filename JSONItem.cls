VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum JSONType
    JSON_Null = 0
    JSON_Bool = 1
    JSON_Number = 2
    JSON_String = 3
    JSON_Array = 4
    JSON_Object = 5
End Enum

Public Property Get Count() As Integer
End Property

Public Property Get Item(Key As String) As JSONItem
End Property

Public Property Get Exists(Key As String) As Boolean
End Property

Public Function Add(Value As JSONItem, Optional Key As String = "") As JSONItem
End Function

Public Function AddArray(Optional Key As String = "") As JSONItem
End Function

Public Function AddObject(Optional Key As String = "") As JSONItem
End Function

Public Function AddNumber(Value As Double, Optional Key As String = "") As JSONItem
End Function

Public Function AddString(Value As String, Optional Key As String = "") As JSONItem
End Function

Public Function AddBoolean(Value As Boolean, Optional Key As String = "") As JSONItem
End Function

Public Function AddNull(Optional Key As String = "") As JSONItem
End Function

Public Sub Remove(Key As String)
End Sub

Public Property Get ItemType() As JSONType
End Property

Public Property Get IsJsArray() As Boolean
End Property

Public Property Get IsJsBool() As Boolean
End Property

Public Property Get IsJsNull() As Boolean
End Property

Public Property Get IsJsNumber() As Boolean
End Property

Public Property Get IsJsObject() As Boolean
End Property

Public Property Get IsJsString() As Boolean
End Property

Public Property Get ToString() As String
End Property

Public Sub SetValue(Value As Variant)
End Sub

Public Function GetValue() As Variant
End Function

Public Property Get Parent() As JSONItem
End Property

Public Property Set Parent(Value As JSONItem)
End Property

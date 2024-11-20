VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get IndentString() As String
    ' empty
End Property

Public Property Let IndentString(Value As String)
    ' empty
End Property

Public Sub ObjectStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectItemStart(Output As JSONOutput, Item As JSONItem, IsFirstItem As Boolean)
    ' empty
End Sub

Public Sub ObjectItemKeyStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectItemKeyWrite(Output As JSONOutput, Item As JSONItem, Key As String)
    ' empty
End Sub

Public Sub ObjectItemKeyEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectItemColonWrite(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectItemValueStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectItemValueWrite(Output As JSONOutput, Item As JSONItem, Value As String)
    ' empty
End Sub

Public Sub ObjectItemEnd(Output As JSONOutput, Item As JSONItem, IsLastItem As Boolean)
    ' empty
End Sub

Public Sub ObjectItemValueEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectItemCommaWrite(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ArrayStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ArrayItemStart(Output As JSONOutput, Item As JSONItem, IsFirstItem As Boolean)
    ' empty
End Sub

Public Sub ArrayItemValueStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ArrayItemValueWrite(Output As JSONOutput, Item As JSONItem, Value As String)
    ' empty
End Sub

Public Sub ArrayItemValueEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ArrayItemCommaWrite(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ArrayItemEnd(Output As JSONOutput, Item As JSONItem, IsLastItem As Boolean)
    ' empty
End Sub

Public Sub ArrayEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub ObjectEmpty(Output As JSONOutput)
    ' empty
End Sub

Public Sub ArrayEmpty(Output As JSONOutput)
    ' empty
End Sub

Public Sub CrLf(Output As JSONOutput)
    ' empty
End Sub


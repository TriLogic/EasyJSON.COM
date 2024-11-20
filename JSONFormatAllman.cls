VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONFormatAllman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements JSONFormat

Dim mIndentString As String
Dim mIndentStack As JSONIndentStack

Private Sub Class_Initialize()
    mIndentString = "  "
    Call Reset
End Sub

Public Sub Reset()
    Set mIndentStack = New JSONIndentStack
    JSONFormat_IndentString = mIndentString
End Sub

Public Property Get JSONFormat_IndentString() As String
    JSONFormat_IndentString = mIndentString
End Property

Public Property Let JSONFormat_IndentString(Value As String)
    mIndentString = Value
    mIndentStack.IndentString = mIndentString
End Property

Public Sub JSONFormat_ObjectStart(Output As JSONOutput, Item As JSONItem)
    mIndentStack.ObjectStart
    If Not Item.Parent Is Nothing Then
        If Item.Parent.IsJsObject Then
            Output.CrLf
            Output.WriteText mIndentStack.OutdentText
        End If
    End If
    Output.WriteText "{"
End Sub

Public Sub JSONFormat_ObjectItemStart(Output As JSONOutput, Item As JSONItem, IsFirstItem As Boolean)
    Output.CrLf
    Output.WriteText mIndentStack.IndentText
End Sub

Public Sub JSONFormat_ObjectItemKeyStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub JSONFormat_ObjectItemKeyWrite(Output As JSONOutput, Item As JSONItem, Key As String)
    Output.WriteText Key
End Sub

Public Sub JSONFormat_ObjectItemKeyEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub JSONFormat_ObjectItemColonWrite(Output As JSONOutput, Item As JSONItem)
    Output.WriteText ": "
End Sub

Public Sub JSONFormat_ObjectItemValueStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub JSONFormat_ObjectItemValueWrite(Output As JSONOutput, Item As JSONItem, Value As String)
    Output.WriteText Value
End Sub

Public Sub JSONFormat_ObjectItemValueEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub JSONFormat_ObjectItemEnd(Output As JSONOutput, Item As JSONItem, IsLastItem As Boolean)
    ' empty
End Sub

Public Sub JSONFormat_ObjectItemCommaWrite(Output As JSONOutput, Item As JSONItem)
    Output.WriteText ", "
End Sub

Public Sub JSONFormat_ObjectEnd(Output As JSONOutput, Item As JSONItem)
    Output.CrLf
    If Not Item.Parent Is Nothing Then
        Output.WriteText mIndentStack.OutdentText
    End If
    Output.WriteText "}"
    mIndentStack.ObjectEnd
End Sub

Public Sub JSONFormat_ArrayStart(Output As JSONOutput, Item As JSONItem)
    mIndentStack.ArrayStart
    If Not Item.Parent Is Nothing Then
        If Item.Parent.IsJsObject Then
            Output.CrLf
            Output.WriteText mIndentStack.OutdentText
        End If
    End If
    Output.WriteText "["
End Sub

Public Sub JSONFormat_ArrayItemStart(Output As JSONOutput, Item As JSONItem, IsFirstItem As Boolean)
    Output.CrLf
    Output.WriteText mIndentStack.IndentText
End Sub

Public Sub JSONFormat_ArrayItemValueStart(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub JSONFormat_ArrayItemValueWrite(Output As JSONOutput, Item As JSONItem, Value As String)
    Output.WriteText Value
End Sub

Public Sub JSONFormat_ArrayItemCommaWrite(Output As JSONOutput, Item As JSONItem)
    Output.WriteText ", "
End Sub

Public Sub JSONFormat_ArrayItemValueEnd(Output As JSONOutput, Item As JSONItem)
    ' empty
End Sub

Public Sub JSONFormat_ArrayItemEnd(Output As JSONOutput, Item As JSONItem, IsLastItem As Boolean)
    ' empty
End Sub

Public Sub JSONFormat_ArrayEnd(Output As JSONOutput, Item As JSONItem)
    Output.CrLf
    If Not Item.Parent Is Nothing Then
        Output.WriteText mIndentStack.OutdentText
    End If
    Output.WriteText "]"
    mIndentStack.ArrayEnd
End Sub

Public Sub JSONFormat_ObjectEmpty(Output As JSONOutput)
    Output.WriteText "{ }"
End Sub

Public Sub JSONFormat_ArrayEmpty(Output As JSONOutput)
    Output.WriteText "[ ]"
End Sub

Public Sub JSONFormat_CrLf(Output As JSONOutput)
    Output.CrLf
End Sub


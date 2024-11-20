VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONIndentStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mIndentStack() As String
Dim mIndentString As String
Dim mStackPointer As Integer

Public Sub Class_Initialize()
    mIndentString = "  "
    Call Reset
End Sub

Public Sub Reset()
    mStackPointer = 0
    ReDim mIndentStack(0)
    mIndentStack(0) = ""
End Sub

Public Property Get IndentString() As String
    IndentString = mIndentString
End Property

Public Property Let IndentString(Value As String)
    mIndentString = Value
End Property

Public Sub ArrayStart()
    mStackPointer = mStackPointer + 1
    If mStackPointer > UBound(mIndentStack) Then
        ReDim Preserve mIndentStack(mStackPointer)
    End If
    mIndentStack(mStackPointer) = OutdentText + mIndentString
End Sub

Public Sub ArrayEnd()
    If mStackPointer = 0 Then
        Err.Raise 66001, , "JSON: Stack underflow"
    End If
    mStackPointer = mStackPointer - 1
End Sub

Public Sub ObjectStart()
    mStackPointer = mStackPointer + 1
    If mStackPointer > UBound(mIndentStack) Then
        ReDim Preserve mIndentStack(mStackPointer)
    End If
    mIndentStack(mStackPointer) = OutdentText + mIndentString
End Sub

Public Sub ObjectEnd()
    If mStackPointer = 0 Then
        Err.Raise 66001, , "JSON: Stack underflow"
    End If
    mStackPointer = mStackPointer - 1
End Sub

Public Property Get StackPointer() As Integer
    StackPointer = mStackPointer
End Property

Public Property Get IndentText() As String
    IndentText = mIndentStack(mStackPointer)
End Property

Public Property Get OutdentText() As String
    If mStackPointer = 0 Then
        OutdentText = ""
    Else
        OutdentText = mIndentStack(mStackPointer - 1)
    End If
End Property


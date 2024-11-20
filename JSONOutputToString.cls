VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONOutputtoString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements JSONOutput

Dim mString As String

Public Property Get StringValue() As String
    StringValue = mString
End Property

Public Property Let StringValue(Value As String)
    mString = Value
End Property

Public Sub JSONOutput_WriteText(Text As String)
    mString = mString + Text
End Sub

Public Sub JSONOutput_WriteLine(Text As String)
    mString = mString + Text + vbCrLf
End Sub

Public Sub JSONOutput_CrLf()
    mString = mString + vbCrLf
End Sub


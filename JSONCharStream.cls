VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONCharStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private mData As String
Private mIndx As Long
Private mLeng As Long

Public Property Get Index() As Long
    Index = mIndx
End Property

Public Property Get Length() As Long
    Length = mIndx
End Property

Public Property Get Buffer() As String
    Buffer = mData
End Property

Public Sub Reset(Value As String)
    mData = Value
    mIndx = 1
    mLeng = Len(mData)
End Sub

Public Function GetC() As Integer
    If mIndx >= 1 And mIndx <= mLeng Then
        GetC = AscW(Mid(mData, mIndx, 1))
        mIndx = mIndx + 1
    Else
        GetC = 0
    End If
End Function

Public Function PeekC() As Integer
    If mIndx >= 1 And mIndx <= mLeng Then
        PeekC = Asc(Mid(mData, mIndx, 1))
    Else
        PeekC = 0
    End If
End Function

Public Sub UngetC(Optional Count As Integer = 1)
    Do While (mIndx > 1 And Count > 1)
        Count = Count - 1
        mIndx = mIndx - 1
    Loop
End Sub


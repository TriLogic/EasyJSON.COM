
Option Explicit

Private mItemArray() As JSONItem
Private mTextArray() As String

Private mIndentText As String

' Object Values
Private mObjectInitSfx As String
Private mObjectOpenPfx As String
Private mObjectOpenSfx As String
Private mObjectKeyPfx As String
Private mObjectKeySfx As String
Private mObjectColonPfx As String
Private mObjectColonSfx As String
Private mObjectValuePfx As String
Private mObjectValueSfx As String
Private mObjectCommaPfx As String
Private mObjectCommaSfx As String
Private mObjectClosePfx As String
Private mObjectCloseSfx As String
Private mObjectFinalPfx As String

Private mObjectEmptyPfx As String
Private mObjectEmptyTxt As String
Private mObjectEmptySfx As String

' Array Values
Private mArrayInitSfx As String
Private mArrayOpenPfx As String
Private mArrayOpenSfx As String
Private mArrayValuePfx As String
Private mArrayArrayPfx As String
Private mArrayArraySfx As String
Private mArrayObjectPfx As String
Private mArrayObjectSfx As String
Private mArrayValueObjF As String
Private mArrayValueSfx As String
Private mArrayCommaPfx As String
Private mArrayCommaSfx As String
Private mArrayClosePfx As String
Private mArrayCloseSfx As String
Private mArrayFinalPfx As String

Private mArrayEmptyPfx As String
Private mArrayEmptyTxt As String
Private mArrayEmptySfx As String

Public Sub Class_Initialize()

    ReDim mItemArray(0)
    Set mItemArray(0) = Nothing
    
    ReDim mTextArray(0)
    Let mTextArray(0) = ""
    
    Call SetFormatKNR

End Sub

Public Property Get IndentText() As String
    IndentText = mIndentText
End Property
Public Property Let IndentText(Value As String)
    mIndentText = Value
End Property

Private Property Get IndentString() As String
    IndentString = mTextArray(UBound(mTextArray))
End Property

Private Property Get OutdentString() As String
    If UBound(mTextArray) > 0 Then
        OutdentString = mTextArray(UBound(mTextArray) - 1)
    Else
        OutdentString = ""
    End If
End Property

Private Sub PushObject(Item As JSONItem)
    ReDim Preserve mItemArray(UBound(mItemArray) + 1)
    Set mItemArray(UBound(mItemArray)) = Item
    ReDim Preserve mTextArray(UBound(mTextArray) + 1)
    Let mTextArray(UBound(mTextArray)) = mTextArray(UBound(mTextArray) - 1) & mIndentText
End Sub
Private Sub PopObject()
    If UBound(mItemArray) > 0 Then
        ReDim Preserve mItemArray(UBound(mItemArray) - 1)
        ReDim Preserve mTextArray(UBound(mTextArray) - 1)
    End If
End Sub

Public Sub ObjectBegin(Item As JSONItem)
    Call PushObject(Item)
End Sub
Public Sub ObjectEnd(Item As JSONItem)
    Call PopObject
End Sub

' Object Style Properties
Public Property Get ObjectInitSfx() As String
    ObjectInitSfx = mObjectInitSfx
End Property
Public Property Let ObjectInitSfx(Value As String)
    mObjectInitSfx = Value
End Property

Public Property Get ObjectOpenPfx() As String
    ObjectOpenPfx = mObjectOpenPfx
End Property
Public Property Let ObjectOpenPfx(Value As String)
    mObjectOpenPfx = Value
End Property
Public Property Get ObjectOpenSfx() As String
    ObjectOpenSfx = mObjectOpenSfx
End Property
Public Property Let ObjectOpenSfx(Value As String)
    mObjectOpenSfx = Value
End Property

' Empty Object
Public Property Get ObjectEmptyPfx() As String
    ObjectEmptyPfx = mObjectEmptyPfx
End Property
Public Property Let ObjectEmptyPfx(Value As String)
    mObjectEmptyPfx = Value
End Property
Public Property Get ObjectEmptyTxt() As String
    ObjectEmptyTxt = mObjectEmptyTxt
End Property
Public Property Let ObjectEmptyTxt(Value As String)
    mObjectEmptyTxt = Value
End Property
Public Property Get ObjectEmptySfx() As String
    ObjectEmptySfx = mObjectEmptySfx
End Property
Public Property Let ObjectEmptySfx(Value As String)
    mObjectEmptySfx = Value
End Property

Public Property Get ObjectColonPfx() As String
    ObjectColonPfx = mObjectColonPfx
End Property
Public Property Let ObjectColonPfx(Value As String)
    mObjectColonPfx = Value
End Property
Public Property Get ObjectColonSfx() As String
    ObjectColonSfx = mObjectColonSfx
End Property
Public Property Let ObjectColonSfx(Value As String)
    mObjectColonSfx = Value
End Property

' First Object In an Array
Public Property Get ArrayArrayPfx() As String
    ArrayArrayPfx = mArrayArrayPfx
End Property
Public Property Let ArrayArrayPfx(Value As String)
    mArrayArrayPfx = Value
End Property
Public Property Get ArrayArraySfx() As String
    ArrayArraySfx = mArrayArraySfx
End Property
Public Property Let ArrayArraySfx(Value As String)
    mArrayArraySfx = Value
End Property

' First Object In an Array
Public Property Get ArrayObjectPfx() As String
    ArrayObjectPfx = mArrayObjectPfx
End Property
Public Property Let ArrayObjectPfx(Value As String)
    mArrayObjectPfx = Value
End Property
Public Property Get ArrayObjectSfx() As String
    ArrayObjectSfx = mArrayObjectSfx
End Property
Public Property Let ArrayObjectSfx(Value As String)
    mArrayObjectSfx = Value
End Property


Public Property Get ObjectValuePfx() As String
    ObjectValuePfx = mObjectValuePfx
End Property
Public Property Let ObjectValuePfx(Value As String)
    mObjectValuePfx = Value
End Property
Public Property Get ObjectValueSfx() As String
    ObjectValueSfx = mObjectValueSfx
End Property
Public Property Let ObjectValueSfx(Value As String)
    mObjectValueSfx = Value
End Property

Public Property Get ObjectCommaPfx() As String
    ObjectCommaPfx = mObjectCommaPfx
End Property
Public Property Let ObjectCommaPfx(Value As String)
    mObjectCommaPfx = Value
End Property
Public Property Get ObjectCommaSfx() As String
    ObjectCommaSfx = mObjectCommaSfx
End Property
Public Property Let ObjectCommaSfx(Value As String)
    mObjectCommaSfx = Value
End Property

Public Property Get ObjectFinalPfx() As String
    ObjectFinalPfx = mObjectFinalPfx
End Property
Public Property Let ObjectFinalPfx(Value As String)
    mObjectFinalPfx = Value
End Property

' Methods and Properties
Public Sub ArrayBegin(Item As JSONItem)
    Call PushObject(Item)
End Sub
Public Sub ArrayEnd(Item As JSONItem)
    Call PopObject
End Sub

' Array Style Properties
Public Property Get ArrayInitSfx() As String
    ArrayInitSfx = mArrayInitSfx
End Property
Public Property Let ArrayInitSfx(Value As String)
    mArrayInitSfx = Value
End Property

Public Property Get ArrayOpenPfx() As String
    ArrayOpenPfx = mArrayOpenPfx
End Property
Public Property Let ArrayOpenPfx(Value As String)
    mArrayOpenPfx = Value
End Property
Public Property Get ArrayOpenSfx() As String
    ArrayOpenSfx = mArrayOpenSfx
End Property
Public Property Let ArrayOpenSfx(Value As String)
    mArrayOpenSfx = Value
End Property

' Empty Array
Public Property Get ArrayEmptyPfx() As String
    ArrayEmptyPfx = mArrayEmptyPfx
End Property
Public Property Let ArrayEmptyPfx(Value As String)
    mArrayEmptyPfx = Value
End Property
Public Property Get ArrayEmptyTxt() As String
    ArrayEmptyTxt = mArrayEmptyTxt
End Property
Public Property Let ArrayEmptyTxt(Value As String)
    mArrayEmptyTxt = Value
End Property
Public Property Get ArrayEmptySfx() As String
    ArrayEmptySfx = mArrayEmptySfx
End Property
Public Property Let ArrayEmptySfx(Value As String)
    mArrayEmptySfx = Value
End Property

Public Property Get ArrayValuePfx() As String
    ArrayValuePfx = mArrayValuePfx
End Property
Public Property Let ArrayValuePfx(Value As String)
    mArrayValuePfx = Value
End Property
Public Property Get ArrayValueSfx() As String
    ArrayValueSfx = mArrayValueSfx
End Property
Public Property Let ArrayValueSfx(Value As String)
    mArrayValueSfx = Value
End Property

Public Property Get ArrayCommaPfx() As String
    ArrayCommaPfx = mArrayCommaPfx
End Property
Public Property Let ArrayCommaPfx(Value As String)
    mArrayCommaPfx = Value
End Property
Public Property Get ArrayCommaSfx() As String
    ArrayCommaSfx = mArrayCommaSfx
End Property
Public Property Let ArrayCommaSfx(Value As String)
    mArrayCommaSfx = Value
End Property

Public Property Get ArrayFinalPfx() As String
    ArrayFinalPfx = mArrayFinalPfx
End Property
Public Property Let ArrayFinalPfx(Value As String)
    mArrayFinalPfx = Value
End Property

Public Function GetExpand(Tpl As String) As String
    Dim Txt As String
    
    If InStr(1, Tpl, "%") < 1 Then
        GetExpand = Tpl
        Exit Function
    End If
    
    Txt = Replace(Tpl, "%crlf%", vbCrLf)
    Txt = Replace(Txt, "%cr%", vbCr)
    Txt = Replace(Txt, "%lf%", vbLf)
    Txt = Replace(Txt, "%tab%", vbTab)
    Txt = Replace(Txt, "%>0%", OutdentString)
    Txt = Replace(Txt, "%>1%", IndentString)
    Txt = Replace(Txt, "%>2%", IndentString + mIndentText)
    Txt = Replace(Txt, "%>+%", mIndentText)
    
    GetExpand = Txt
    
End Function

Public Function GetWrapped(Txt As String, TplPfx As String, TplSfx As String) As String

    Dim Prefix As String
    Dim Suffix As String
    Dim Result As String
    
    Prefix = GetExpand(TplPfx)
    Suffix = GetExpand(TplSfx)
    
    Result = Prefix + Txt + Suffix
    GetWrapped = Result

End Function

Public Function ToString(Item As JSONItem) As String
    ToString = ToStringInternal(Item, True, True)
End Function

Private Function ToStringInternal(Item As JSONItem, IsFirst As Boolean, IsBase As Boolean) As String

    If Item.IsJsArray Then
        ToStringInternal = ToStringArray(Item, IsFirst, IsBase)
    ElseIf Item.IsJsObject Then
        ToStringInternal = ToStringObject(Item, IsFirst, IsBase)
    Else
        ToStringInternal = Item.ToString()
    End If

End Function

Private Function ToStringObject(Item As JSONItem, IsFirst As Boolean, IsBase As Boolean) As String

    Dim ObjBeg As String
    Dim ObjEnd As String
    Dim Result As String
    Dim Key As Variant
    Dim OutCount As Integer
    
    ' Format an empty that is base
    If Item.Count = 0 Then
        Result = "{" & mObjectEmptyTxt & "}"
        If IsBase Then
            ToStringObject = Result
        Else
            If IsFirst Or Item.Parent.IsJsArray Then
                ToStringObject = GetWrapped(Result, mArrayValuePfx, mArrayValueSfx)
            Else
                ToStringObject = GetWrapped(Result, mObjectEmptyPfx, mObjectEmptySfx)
            End If
        End If
        Exit Function
    End If
    
    If IsBase Then
        ObjBeg = GetWrapped("{", "", mObjectInitSfx)
        ObjEnd = GetWrapped("}", mObjectFinalPfx, "")
    Else
        If IsFirst And Item.Parent.IsJsArray Then
            ObjBeg = GetWrapped("{", mArrayObjectPfx, mArrayObjectSfx)
        Else
            ObjBeg = GetWrapped("{", mObjectOpenPfx, mObjectOpenSfx)
        End If
        ObjEnd = GetWrapped("}", mObjectClosePfx, mObjectCloseSfx)
    End If
        
    Call ObjectBegin(Item)
    
    Dim D As Dictionary
    Dim I As JSONItem
    Set D = Item.GetValue()
    
    For Each Key In D.Keys
    
        Set I = D(Key)
    
        OutCount = OutCount + 1
        If OutCount > 1 Then
            Result = Result & GetWrapped(",", mObjectCommaPfx, mObjectCommaSfx)
        End If
        
        Result = Result & GetWrapped("""" & Key & """", mObjectKeyPfx, mObjectKeySfx)
        Result = Result & GetWrapped(":", mObjectColonPfx, mObjectColonSfx)
        If (Item.IsJsArray Or Item.IsJsObject) And (Item.Count = 0) Then
            Result = Result & ToStringInternal(I, Item.Count = 0, False)
        Else
            Result = Result & GetWrapped(ToStringInternal(I, Item.Count = 0, False), mObjectValuePfx, mObjectValueSfx)
        End If
        
    Next
    
    Call ObjectEnd(Item)
    
    Result = ObjBeg & Result & ObjEnd
    
    ToStringObject = Result
    
End Function

Private Function ToStringArray(Item As JSONItem, IsFirst As Boolean, IsBase As Boolean) As String
    
    Dim ObjBeg As String
    Dim ObjEnd As String
    Dim Result As String
    Dim Key As Variant
    Dim OutCount As Integer

    ' Format an Empty Array that is base
    If Item.Count = 0 Then
        Result = "[" & mArrayEmptyTxt & "]"
        If IsBase Then
            ToStringArray = Result
        Else
            If IsFirst Or Item.Parent.IsJsArray Then
                ToStringArray = GetWrapped(Result, mArrayValuePfx, mArrayValueSfx)
            Else
                ToStringArray = GetWrapped(Result, mArrayEmptyPfx, mArrayEmptySfx)
            End If
        End If
        Exit Function
    End If

    If IsBase Then
        ' Item is the Array child that is the base object being written.
        ObjBeg = GetWrapped("[", "", mArrayInitSfx)
        ObjEnd = GetWrapped("]", mArrayFinalPfx, "")
    Else
        ' Item is the Array child of another Array or Object
        ' If we're the FIRST Object being written that's different
        If IsFirst Then
            ObjBeg = GetWrapped("[", mArrayArrayPfx, mArrayArraySfx)
        Else
            ObjBeg = GetWrapped("[", mArrayOpenPfx, mArrayOpenSfx)
        End If
        ObjEnd = GetWrapped("]", mArrayClosePfx, mArrayCloseSfx)
    End If
        
    Call ArrayBegin(Item)

    Dim D As Variant
    Dim I As JSONItem
    D = Item.GetValue

    For Key = 1 To Item.Count
    
        Set I = D(Key)
        OutCount = OutCount + 1
        
        If OutCount > 1 Then
            Result = Result & GetWrapped(",", mArrayCommaPfx, mArrayCommaSfx)
        End If
        
        ' Was Item.IsJsArray or Item.IsJsObject
        If (I.IsJsArray Or I.IsJsObject) Then
            Result = Result & ToStringInternal(I, OutCount = 1, False)
        Else
            Result = Result & GetWrapped(ToStringInternal(I, OutCount = 1, False), mArrayValuePfx, mArrayValueSfx)
        End If

    Next Key
    
    Call ArrayEnd(Item)
    
    Result = ObjBeg & Result & ObjEnd
    
    ToStringArray = Result
    
End Function

Public Sub SetFormatAllman(Optional IndentTxt As String = vbTab, Optional ExpandEmpty As Boolean = False, Optional BlankBeforeColon As Boolean = False)

    mIndentText = IndentTxt

    ' Object Formatting
    mObjectInitSfx = ""
    mObjectOpenPfx = "%crlf%%>1%"
    mObjectOpenSfx = ""
    mObjectKeyPfx = "%crlf%%>1%"
    mObjectKeySfx = ""
    mObjectColonPfx = IIf(BlankBeforeColon, " ", "")
    mObjectColonSfx = " "
    mObjectValuePfx = ""
    mObjectValueSfx = ""
    mObjectCommaPfx = ""
    mObjectCommaSfx = ""
    mObjectClosePfx = "%crlf%%>1%"
    mObjectCloseSfx = ""
    mObjectFinalPfx = "%crlf%"
    
    mObjectEmptyPfx = ""
    mObjectEmptyTxt = IIf(ExpandEmpty, " ", "")
    mObjectEmptySfx = ""
    
    ' Array Values
    mArrayInitSfx = ""
    mArrayOpenPfx = "%crlf%%>1%"
    mArrayOpenSfx = ""
    mArrayArrayPfx = "%crlf%%>1%"
    mArrayArraySfx = ""
    mArrayObjectPfx = "%crlf%%>1%"
    mArrayObjectSfx = ""
    mArrayValuePfx = "%crlf%%>1%"
    mArrayValueSfx = ""
    mArrayCommaPfx = ""
    mArrayCommaSfx = ""
    mArrayClosePfx = "%crlf%%>1%"
    mArrayCloseSfx = ""
    mArrayFinalPfx = "%crlf%"
    
    mArrayEmptyPfx = ""
    mArrayEmptyTxt = IIf(ExpandEmpty, " ", "")
    mArrayEmptySfx = ""

End Sub

Public Sub SetFormatWhitesmith(Optional IndentTxt As String = vbTab, Optional ExpandEmpty As Boolean = False, Optional BlankBeforeColon As Boolean = False)

    mIndentText = IndentTxt

    ' Object Formatting
    mObjectInitSfx = ""
    mObjectOpenPfx = "%crlf%%>2%"
    mObjectOpenSfx = ""
    mObjectKeyPfx = "%crlf%%>1%"
    mObjectKeySfx = ""
    mObjectColonPfx = IIf(BlankBeforeColon, " ", "")
    mObjectColonSfx = " "
    mObjectValuePfx = ""
    mObjectValueSfx = ""
    mObjectCommaPfx = ""
    mObjectCommaSfx = ""
    mObjectClosePfx = "%crlf%%>2%"
    mObjectCloseSfx = ""
    mObjectFinalPfx = "%crlf%"
    
    mObjectEmptyPfx = ""
    mObjectEmptyTxt = IIf(ExpandEmpty, " ", "")
    mObjectEmptySfx = ""
    
    ' Array Values
    mArrayInitSfx = ""
    mArrayOpenPfx = "%crlf%%>2%"
    mArrayOpenSfx = ""
    mArrayArrayPfx = "%crlf%%>2%"
    mArrayArraySfx = ""
    mArrayObjectPfx = "%crlf%%>2%"
    mArrayObjectSfx = ""
    mArrayValuePfx = "%crlf%%>1%"
    mArrayValueSfx = ""
    mArrayCommaPfx = ""
    mArrayCommaSfx = ""
    mArrayClosePfx = "%crlf%%>2%"
    mArrayCloseSfx = ""
    mArrayFinalPfx = "%crlf%"
    
    mArrayEmptyPfx = ""
    mArrayEmptyTxt = IIf(ExpandEmpty, " ", "")
    mArrayEmptySfx = ""

End Sub

Public Sub SetFormatKNR(Optional IndentTxt As String = vbTab, Optional ExpandEmpty As Boolean = False, Optional OutdentClose As Boolean = False)

    mIndentText = IndentTxt

    ' Object Formatting
    mObjectInitSfx = ""
    mObjectOpenPfx = ""
    mObjectOpenSfx = ""
    mObjectKeyPfx = "%crlf%%>1%"
    mObjectKeySfx = ""
    mObjectColonPfx = ""
    mObjectColonSfx = " "
    mObjectValuePfx = ""
    mObjectValueSfx = ""
    mObjectCommaPfx = ""
    mObjectCommaSfx = ""
    mObjectClosePfx = IIf(OutdentClose, "%crlf%%>1%", "%crlf%%>2%")
    mObjectCloseSfx = ""
    mObjectFinalPfx = "%crlf%"
    
    mObjectEmptyPfx = ""
    mObjectEmptyTxt = IIf(ExpandEmpty, " ", "")
    mObjectEmptySfx = ""
    
    ' Array Values
    mArrayInitSfx = ""
    mArrayOpenPfx = ""
    mArrayOpenSfx = ""
    mArrayArrayPfx = " "
    mArrayArraySfx = ""
    mArrayObjectPfx = " "
    mArrayObjectSfx = ""
    mArrayValuePfx = "%crlf%%>1%"
    mArrayValueSfx = ""
    mArrayCommaPfx = ""
    mArrayCommaSfx = " "
    mArrayClosePfx = IIf(OutdentClose, "%crlf%%>1%", "%crlf%%>2%")
    mArrayCloseSfx = ""
    mArrayFinalPfx = "%crlf%"
    
    mArrayEmptyPfx = ""
    mArrayEmptyTxt = IIf(ExpandEmpty, " ", "")
    mArrayEmptySfx = ""

End Sub

Public Sub SetFormatLinear(Optional Expanded As Boolean = True)

    mIndentText = ""

    ' Object Formatting
    mObjectInitSfx = IIf(Expanded, " ", "")
    mObjectOpenPfx = ""
    mObjectOpenSfx = IIf(Expanded, " ", "")
    mObjectKeyPfx = ""
    mObjectKeySfx = ""
    mObjectColonPfx = IIf(Expanded, " ", "")
    mObjectColonSfx = IIf(Expanded, " ", "")
    mObjectValuePfx = ""
    mObjectValueSfx = ""
    mObjectCommaPfx = ""
    mObjectCommaSfx = IIf(Expanded, " ", "")
    mObjectClosePfx = IIf(Expanded, " ", "")
    mObjectCloseSfx = ""
    mObjectFinalPfx = ""
    
    mObjectEmptyPfx = ""
    mObjectEmptyTxt = IIf(Expanded, " ", "")
    mObjectEmptySfx = ""
    
    ' Array Values
    mArrayInitSfx = IIf(Expanded, " ", "")
    mArrayOpenPfx = ""
    mArrayOpenSfx = IIf(Expanded, " ", "")
    mArrayArrayPfx = IIf(Expanded, " ", "")
    mArrayArraySfx = ""
    mArrayObjectPfx = IIf(Expanded, " ", "")
    mArrayObjectSfx = ""
    mArrayValuePfx = ""
    mArrayValueSfx = ""
    mArrayCommaPfx = ""
    mArrayCommaSfx = IIf(Expanded, " ", "")
    mArrayClosePfx = IIf(Expanded, " ", "")
    mArrayCloseSfx = ""
    mArrayFinalPfx = IIf(Expanded, " ", "")
    
    mArrayEmptyPfx = ""
    mArrayEmptyTxt = IIf(Expanded, " ", "")
    mArrayEmptySfx = ""

End Sub

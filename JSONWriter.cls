VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSONWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mObjectType() As JSONType
Dim mCountStack() As Long
Dim mStackPointer As Integer
Dim mOutput As JSONOutput
Dim mFormat As JSONFormat

Private Sub Class_Initialize()
    Call Reset
End Sub

Public Sub Reset()
    ReDim mCountStack(0)
    mCountStack(0) = 0
    ReDim mObjectType(0)
    mObjectType(0) = JSON_Null
    mStackPointer = 0
    Set mFormat = New JSONFormatCompact
    Set mOutput = New JSONOutputtoString
End Sub

Public Property Get ItemCount() As Long
    ItemCount = mCountStack(mStackPointer)
End Property

Public Property Get StackPointer() As Integer
    StackPointer = mStackPointer
End Property

Public Property Get IsArrayTOS() As Boolean
    IsArrayTOS = (mObjectType(mStackPointer) = JSON_Array)
End Property

Public Property Get IsObjectTOS() As Boolean
    IsObjectTOS = (mObjectType(mStackPointer) = JSON_Object)
End Property

Public Property Get IsContainerTOS() As Boolean
    IsContainerTOS = (IsArrayTOS Or IsObjectTOS)
End Property

Public Property Get Output() As JSONOutput
    Set Output = mOutput
End Property

Public Property Set Output(Value As JSONOutput)
    Set mOutput = Value
End Property

Public Property Get Format() As JSONFormat
    Set Format = mFormat
End Property

Public Property Set Format(Value As JSONFormat)
    Set mFormat = Value
End Property

Public Sub WriteItem(Item As JSONItem)

    Select Case Item.ItemType
    Case JSON_Null
        mOutput.WriteText Item.ToString()
    Case JSON_Bool
        mOutput.WriteText Item.ToString()
    Case JSON_Number
        mOutput.WriteText Item.ToString()
    Case JSON_String
        mOutput.WriteText """" & ToEscapedString(Item.ToString()) & """"
    Case JSON_Array
        WriteItemArray Item
    Case JSON_Object
        WriteItemObject Item
    End Select

End Sub

Public Sub WriteItemArray(Item As JSONItem)

    Dim WasContainer As Boolean

    If Item.Count = 0 Then
        mFormat.ArrayEmpty mOutput
        Exit Sub
    End If

    mFormat.ArrayStart mOutput, Item
    
    ' A is an array of Variant(JSItem)
    Dim A As Variant
    Let A = Item.GetValue()
    
    Dim Idx As Long
    Dim C As JSONItem
    
    For Idx = 1 To UBound(A)
        WasContainer = False
    
        ' Retrieve the child item
        Set C = A(Idx)
        
        mFormat.ArrayItemStart mOutput, C, Idx = 1
        mFormat.ArrayItemValueStart mOutput, C
        
        Select Case C.ItemType
        Case JSON_Null
            mOutput.WriteText C.ToString()
        Case JSON_Bool
            mOutput.WriteText C.ToString()
        Case JSON_Number
            mOutput.WriteText C.ToString()
        Case JSON_String
            mOutput.WriteText C.ToString()
        Case JSON_Array
            WriteItemArray C
            WasContainer = True
        Case JSON_Object
            WriteItemObject C
            WasContainer = True
        End Select
        
        mFormat.ArrayItemValueEnd mOutput, C
        mFormat.ArrayItemEnd mOutput, C, Idx = UBound(A)
        
        ' If this is the last item then no comma is written
        If Idx < UBound(A) Then
            mFormat.ArrayItemCommaWrite mOutput, C
        End If

    Next
    
    mFormat.ArrayEnd mOutput, Item

End Sub

Public Sub WriteItemObject(Item As JSONItem)

    Dim WasContainer As Boolean
    Dim ItemCount As Long

    If Item.Count = 0 Then
        mFormat.ObjectEmpty mOutput
        Exit Sub
    End If

    mFormat.ObjectStart mOutput, Item
    
    Dim D As Dictionary
    Dim V As Variant
    Dim C As JSONItem
    Dim K As String
    
    Set D = Item.GetValue()
    
    ItemCount = 0
    For Each V In D.Keys
    
        WasContainer = False
        ItemCount = ItemCount + 1
    
        ' Retrieve the child item
        K = CStr(V)
        Set C = D(K)
    
        If K = "Array" Then
            Debug.Print "Found 'Array'"
        End If
            
    
        mFormat.ObjectItemStart mOutput, C, ItemCount = 1
        mFormat.ObjectItemKeyStart mOutput, C
        mFormat.ObjectItemKeyWrite mOutput, C, """" & ToEscapedString(K) & """"
        mFormat.ObjectItemKeyEnd mOutput, C
        mFormat.ObjectItemColonWrite mOutput, C
        mFormat.ObjectItemValueStart mOutput, C
        
        Select Case C.ItemType
        Case JSON_Null
            mOutput.WriteText C.ToString
        Case JSON_Bool
            mOutput.WriteText C.ToString
        Case JSON_Number
            mOutput.WriteText C.ToString
        Case JSON_String
            mOutput.WriteText C.ToString
        Case JSON_Array
            WriteItemArray C
            WasContainer = True
        Case JSON_Object
            WriteItemObject C
            WasContainer = True
        End Select

        mFormat.ObjectItemValueEnd mOutput, C
        mFormat.ObjectItemEnd mOutput, C, ItemCount = D.Count
        
        If ItemCount < D.Count Then
            mFormat.ObjectItemCommaWrite mOutput, C
        End If

    Next
    
    mFormat.ObjectEnd mOutput, Item

End Sub

Public Sub SetFormatCompact()
    Set Format = New JSONFormatCompact
End Sub

Public Sub SetFormatExpanded()
    Set Format = New JSONFormatExpanded
End Sub

Public Sub SetFormatAllman()
    Set Format = New JSONFormatAllman
End Sub

Public Sub SetFormatKNR()
    Set Format = New JSONFormatKNR
End Sub

Public Sub SetFormatWhitesmith()
    Set Format = New JSONFormatWhitesmith
End Sub

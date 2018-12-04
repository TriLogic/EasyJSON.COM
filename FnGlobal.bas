Attribute VB_Name = "FnGlobal"
Option Explicit

Public Function Replace(ByVal strSource As String, ByVal strWhat As String, ByVal strWith As String) As String

    Dim intstrWhatLength As Integer
    Dim intReplaceLength As Integer
    Dim intStart As Integer
    
    intstrWhatLength = Len(strWhat)
    
    If intstrWhatLength = 0 Then
        Replace = strSource
        Exit Function
    End If
    
    intReplaceLength = Len(strWith)
    intStart = InStr(1, strSource, strWhat)
    
    Do While intStart > 0
        strSource = Left(strSource, intStart - 1) + strWith + Right(strSource, Len(strSource) - (intStart + intstrWhatLength - 1))
        intStart = InStr(intStart + intReplaceLength, strSource, strWhat)
    Loop
    
    Replace = strSource

End Function


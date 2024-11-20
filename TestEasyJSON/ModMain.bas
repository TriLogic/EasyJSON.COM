Attribute VB_Name = "ModMain"
Option Explicit

Sub Main()

    Dim Result As String

    Call TestTokenizer
    'Call TestReader
    'Call TestFormatter

End Sub

Public Sub TestReader()

    ' Build the JSON Object
    Dim oOutJSON As JSONItem
    Dim sOutJSON As String
    
    ' Build JSON and Convert to string
    Set oOutJSON = BuildMegaObject
    sOutJSON = oOutJSON.ToString()
    
    ' Retrieve the JSON String
    Dim oInpJSON As JSONItem
    Dim sInpJSON As String
    sInpJSON = oOutJSON.ToString
    
    ' Read the JSON back to a JSON object
    Dim oRead As New JSONReader
    Set oInpJSON = oRead.GetObject(sOutJSON)
    
    ' Write the input JSON back to a string
    sInpJSON = oInpJSON.ToString()
    
    Debug.Print "TestReader: " & IIf(sOutJSON = sInpJSON, "Passed", "Failed")
    
End Sub

Public Sub TestTokenizer()

    ' Define variables
    Dim tokens As New JSONTokenizer
    Dim oJSON As JSONItem
    Dim sJOSN As String
    
    ' Build a JSON Object
    Set oJSON = BuildSuperObject()
    
    
    ' Create a writer and formatting
    Dim oWriter As New JSONWriter
    Dim oOutput As New JSONOutputtoString
    Set oWriter.Output = oOutput
    
    ' Set formatting
    oWriter.SetFormatAllman
    
    ' Write to the writer
    oWriter.WriteItem oJSON
    
    Debug.Print oOutput.StringValue
    
    ' Initialize the tokenizer
    tokens.Reset oOutput.StringValue
    
    ' Read and print the tokens
    Do While tokens.GetToken()
        Debug.Print "TOKEN=""" & tokens.GetEscapedString(tokens.TokenValue) & """"
        If tokens.TokenType = TK_EOF Then
            Exit Do
        End If
    Loop

End Sub

Public Sub TestFormatter()

    ' Define variables
    Dim oWriter As JSONWriter
    Dim oOutput As JSONOutputtoString
    Dim oJSON As JSONItem
    
    ' Build a JSON Object Tree
    Set oJSON = BuildObjectOfAll()
    'Set oJson = BuildArrayOfAll()
     
    ' Create the writer and formatting
    Set oWriter = New JSONWriter
    Set oOutput = New JSONOutputtoString
    Set oWriter.Output = oOutput
    
    ' Set formatting
    oWriter.SetFormatAllman
    'oWriter.SetFormatWhitesmith
    'oWriter.SetFormatKNR
    'oWriter.SetFormatCompact
    'oWriter.SetFormatExpanded
    
    ' For Testing purposes only
    'oWriter.Format.IndentString = "..."
    
    ' Write to the writer
    oWriter.WriteItem oJSON
    
    ' Output the string
    Debug.Print oOutput.StringValue

End Sub



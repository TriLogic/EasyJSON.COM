Attribute VB_Name = "ModBuilder"
Option Explicit

Public Function BuildObjectOfNull() As JSONItem

    Dim O As JSONItem
    Set O = New JSONObject
    
    O.AddNull "Null_1"
    O.AddNull "Null_2"
    
    Set BuildObjectOfNull = O
    
End Function

Public Function BuildObjectOfBoolean() As JSONItem

    Dim O As JSONItem
    Set O = New JSONObject

    O.AddBoolean True, "True"
    O.AddBoolean False, "False"

    Set BuildObjectOfBoolean = O

End Function

Public Function BuildObjectOfNumber() As JSONItem

    Dim O As JSONItem
    Set O = New JSONObject
    
    O.AddNumber 0, "zero"
    O.AddNumber 1, "one"
    O.AddNumber -3, "negative"
    O.AddNumber 0.005, "decimal"
    O.AddNumber -5.79E-32, "scientific"
    
    Set BuildObjectOfNumber = O

End Function

Public Function BuildObjectOfString() As JSONItem

    Dim O As JSONItem
    Set O = New JSONObject
    
    O.AddString "Hello", "String1"
    O.AddString "World", "String2"

    Set BuildObjectOfString = O

End Function

Public Function BuildObjectOfSpecialString() As JSONItem
    
    Dim O As JSONItem
    Set O = New JSONObject
    
    O.AddString "", "empty"
    O.AddString vbCr, "vbCr"
    O.AddString vbLf, "vbLf"
    O.AddString vbTab, "vbTab"
    O.AddString Chr(8), "BackSpace"
    O.AddString "\", "BackSlash"
    O.AddString "/", "Slash"
    O.AddString """", "DoubleQote"
    O.AddString ChrW(&H110), "Unicode"
    
    Set BuildObjectOfSpecialString = O

End Function

Public Function BuildObjectOfEmptyObject() As JSONItem

    Dim O As JSONItem
    Set O = New JSONObject

    O.AddObject "Empty1"
    O.AddObject "Empty2"
    O.AddObject "Empty3"

    Set BuildObjectOfEmptyObject = O

End Function

Public Function BuildObjectOfEmptyArray() As JSONItem

    Dim O As JSONItem
    Set O = New JSONObject

    O.AddArray "Empty1"
    O.AddArray "Empty2"
    O.AddArray "Empty3"

    Set BuildObjectOfEmptyArray = O

End Function

Public Function BuildObjectOfAll() As JSONItem

    Dim A As JSONItem
    Set A = New JSONObject

    A.AddNull "Null"
    A.AddBoolean True, "Boolean"
    A.AddNumber "123", "Number"
    A.AddString "Hello World!", "String"
    
    A.Add BuildArrayOfNull(), "ArrayOfNull"
    A.Add BuildArrayOfBoolean(), "ArrayOfBoolean"
    A.Add BuildArrayOfNumber(), "ArrayOfNumber"
    A.Add BuildArrayOfString(), "ArrayOfString"
    A.Add BuildArrayOfEmptyArray(), "ArrayOfEmptyArray"
    A.Add BuildArrayOfEmptyObject(), "ArrayOfEmptyObject"
    
    A.Add BuildObjectOfNull(), "ObjectOfNull"
    A.Add BuildObjectOfBoolean(), "ObjectOfBoolean"
    A.Add BuildObjectOfNumber(), "ObjectOfNumber"
    A.Add BuildObjectOfString(), "ObjectOfString"
    A.Add BuildObjectOfEmptyArray(), "ObjectOfEmptyArray"
    A.Add BuildObjectOfEmptyObject(), "ObjectOfEmptyObject"
    A.Add BuildObjectOfSpecialString(), "ObjectOfSpecialString"

    Set BuildObjectOfAll = A

End Function

Public Function BuildSuperObject() As JSONItem

    Dim A As JSONItem
    Set A = New JSONObject
    
    A.Add BuildArrayOfAll(), "ArrayOfAll"
    A.Add BuildObjectOfAll(), "ObjectOfAll"
    
    Set BuildSuperObject = A

End Function

Public Function BuildMegaObject() As JSONItem

    Dim A As JSONItem
    Set A = New JSONObject
    
    A.Add BuildSuperArray(), "SuperArray"
    A.Add BuildSuperObject(), "SuperObject"
    
    Set BuildMegaObject = A

End Function

Public Function BuildArrayOfNull() As JSONItem

    Dim O As JSONItem
    Set O = New JSONArray
    
    O.AddNull
    O.AddNull
    
    Set BuildArrayOfNull = O
    
End Function

Public Function BuildArrayOfBoolean() As JSONItem

    Dim O As JSONItem
    Set O = New JSONArray

    O.AddBoolean True
    O.AddBoolean False

    Set BuildArrayOfBoolean = O

End Function

Public Function BuildArrayOfNumber() As JSONItem

    Dim O As JSONItem
    Set O = New JSONArray
    
    O.AddNumber 0
    O.AddNumber 1
    O.AddNumber -3
    O.AddNumber 0.005
    O.AddNumber -5.79E-32
    
    Set BuildArrayOfNumber = O

End Function

Public Function BuildArrayOfString() As JSONItem

    Dim O As JSONItem
    Set O = New JSONArray

    O.AddString "Hello"
    O.AddString "World"

    Set BuildArrayOfString = O

End Function

Public Function BuildArrayOfEmptyArray() As JSONItem

    Dim A As JSONItem
    Set A = New JSONArray
    
    A.AddArray
    A.AddArray
    A.AddArray

    Set BuildArrayOfEmptyArray = A

End Function

Public Function BuildArrayOfEmptyObject() As JSONItem

    Dim A As JSONItem
    Set A = New JSONArray
    
    A.AddObject
    A.AddObject
    A.AddObject

    Set BuildArrayOfEmptyObject = A

End Function

Public Function BuildArrayOfAll() As JSONItem

    Dim A As JSONItem
    Set A = New JSONArray

    A.AddNull
    A.AddBoolean True
    A.AddNumber "123"
    A.AddString "Hello World!"
    
    A.Add BuildArrayOfNull()
    A.Add BuildArrayOfBoolean()
    A.Add BuildArrayOfNumber()
    A.Add BuildArrayOfString()
    A.Add BuildArrayOfEmptyArray()
    A.Add BuildArrayOfEmptyObject()
    
    A.Add BuildObjectOfNull()
    A.Add BuildObjectOfBoolean()
    A.Add BuildObjectOfNumber()
    A.Add BuildObjectOfString()
    A.Add BuildObjectOfEmptyArray()
    A.Add BuildObjectOfEmptyObject()
    A.Add BuildObjectOfSpecialString()

    Set BuildArrayOfAll = A

End Function

Public Function BuildSuperArray() As JSONItem

    Dim A As JSONItem
    Set A = New JSONArray
    
    A.Add BuildArrayOfAll()
    A.Add BuildObjectOfAll()
    
    Set BuildSuperArray = A

End Function

Public Function BuildMegaArray() As JSONItem

    Dim A As JSONItem
    Set A = New JSONArray
    
    A.Add BuildSuperArray()
    A.Add BuildSuperObject()
    
    Set BuildMegaArray = A

End Function


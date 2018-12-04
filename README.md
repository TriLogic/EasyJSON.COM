# EasyJSON.COM
A simple library for reading, writing and creating JSON in VB5, VB6, and VBA

This library provides a simple and straightforward way to create, read and write JSON data.  It is language compatible 
with older Visual Basic environments VB5, VB6 as well as VBA (MS Word, MS Excel, MS Access).  I believe it to be 
fully compatible with the latest JSON specification.

Creating JSON:

JSON data is created by creating an instance of JSONObject and then adding values including:
  JSONObject, JSONArray, JSONString, JSONBoolean, JSONNumber, JSONNull all of which implement the common interface JSONItem.
  
Reading JSON:

JSON data in the form of a string is passed to an instance of the JSONReader which produces a hierarchy of JSON objects.

Writing JSON:

Once a hierachy of objects has been created you can output a JSON string in one of two ways.  
  1) The easiest is simple to call the ToString method of the root object.  This produces a compacted JSON string.
  2) Formatted output functionality is available via the JSONWriter class which provides formatting with options for 
  several popular "C" bracing and indentation styles including: K&R, Allman, Whitesmith as well as Linear compact and expanded.
  
 There is a fuilly functional test application in the TestJSON folder.
  
~Enjoy

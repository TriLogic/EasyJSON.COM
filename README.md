# EasyJSON.COM
A simple library for reading, writing and creating JSON in VB5, VB6, and VBA

This library provides a simple and straightforward way to create, read and write JSON data.  It is language compatible 
with older Visual Basic environments VB5, VB6 as well as VBA (MS Word, MS Excel, MS Access).  I believe it to be 
fully compatible with the latest JSON specification.

Creating JSON:

JSON data is created by creating an instance of **JSONObject** or **JSONArray** and then adding values including:
  **JSONObject**, **JSONArray**, **JSONString**, **JSONBoolean**, **JSONNumber**, **JSONNull** all of which implement
  the common interface **JSONItem**.
  
Reading JSON:

JSON data in the form of a string is passed to an instance of the **JSONReader** which produces a hierarchy of JSON objects.

Writing JSON:

Once a hierachy of objects has been created (or read) you can output a JSON string in one of two ways.  
  1) The easiest is to call the **ToString** method of the root object.  This produces a compacted JSON string.
  2) Formatted output is available via the updated **JSONWriter** class which provides formatting with options for
     several popular "C" bracing and indentation styles including: K&R, Allman, Whitesmith, Expanded and Compact.
  4) A new **JSONOutput** interface with a compatible **JSONOutputToString** object has been defined to simplify writing
     JSON to additional output targets.
  
 There is a fully functional test application in the TestEasyJSON folder.

To Do:
  1) Implement code to replace a item in JSONArray by numeric index key
  2) Implment code to replace an item in JSONObject by string key

~Enjoy

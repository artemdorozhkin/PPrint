Attribute VB_Name = "Examples"
'@Folder "PPrintProject.Examples"
Option Explicit

Public Sub Example()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim c As Collection
    Set c = New Collection

    c.Add "9"
    c.Add 8
    c.Add Array(7, 6, "5")

    d("number") = 1
    d("string") = "c"
    d("array") = Array(1, "2")
    Set d("collection") = c

    PPrint "number:", 1
    PPrint "string:", "Hello, World!"
    PPrint "array:", Array(1, "2")
    PPrint "collection:", c
    PPrint "dictionary:", d
    PPrint "range:", Range("A1:C1")
    PPrint "user class without repr__ method:", New Class1
    PPrint "user class with repr__ method:", New Class2
    PPrint "other objects:", ThisWorkbook, Worksheets, CreateObject("Scripting.FileSystemObject")
End Sub

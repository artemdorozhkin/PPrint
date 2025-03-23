Attribute VB_Name = "PPrintModule"
'@Folder "PPrintProject.src"
Option Explicit

Public Sub PPrint(ParamArray Values() As Variant)
    Dim Result As String

    Dim Item As Variant
    For Each Item In Values
        Result = Result & ToString(Item) & " "
    Next

    If Strings.Len(Result) = 0 Then
        Debug.Print
    Else
        Debug.Print Strings.Left(Result, Strings.Len(Result) - 1)
    End If
End Sub

Private Function ToString(ByRef Item As Variant) As String
    If Information.VarType(Item) = VbVarType.vbString Then
        ToString = GetString(Item)
    ElseIf Information.VarType(Item) = VbVarType.vbDate Then
        ToString = GetDate(Item)
    ElseIf Information.IsObject(Item) Then
        If IsBuiltinObject(Item) Then
            ToString = GetBuiltinObject(Item)
        ElseIf HasReprFunction(Item) Then
            ToString = GetUserDefineObject(Item)
        Else
            ToString = "<object '" & Information.TypeName(Item) & "'>"
        End If
    ElseIf Information.IsArray(Item) Then
        ToString = GetArray(Item)
    Else
        ToString = Conversion.CStr(Item)
    End If
End Function

Private Function IsBuiltinObject(ByRef Item As Variant) As Boolean
    Dim Builtins As String
    Builtins = ";Range;Collection;Dictionary;"
    IsBuiltinObject = Strings.InStr(1, Builtins, ";" & Information.TypeName(Item) & ";")
End Function

Private Function GetBuiltinObject(ByRef Item As Variant) As String
    Dim Name As String
    Name = Information.TypeName(Item)

    If Name = "Range" Then
        GetBuiltinObject = GetRange(Item)
    ElseIf Name = "Collection" Then
        GetBuiltinObject = GetCollection(Item)
    ElseIf Name = "Dictionary" Then
        GetBuiltinObject = GetDictionary(Item)
    End If
End Function

Private Function HasReprFunction(ByRef Item As Variant) As Boolean
    On Error Resume Next
    Item.Repr__
    HasReprFunction = Information.Err().Number = 0
End Function

Private Function GetUserDefineObject(ByRef Item As Variant) As String
    GetUserDefineObject = Item.Repr__()
End Function

Private Function GetArray(ByRef Elements As Variant) As String
    GetArray = "[" & GetIterable(Elements) & "]"
End Function

Private Function GetRange(ByRef Range As Variant) As String
    GetRange = "Range<" & Range.Address & ">" & GetArray(Range.Value)
End Function

Private Function GetCollection(ByRef Elements As Variant) As String
    GetCollection = "(" & GetIterable(Elements) & ")"
End Function

Private Function GetDictionary(ByRef Elements As Variant) As String
    Dim Result As String
    Dim Key As Variant
    For Each Key In Elements
        Result = Result & GetString(Key) & ": "
        Result = Result & ToString(Elements(Key)) & ", "
    Next

    If Strings.Len(Result) < 2 Then
        GetDictionary = "{}"
    Else
        GetDictionary = "{" & Strings.Left(Result, Strings.Len(Result) - 2) & "}"
    End If
End Function

Private Function GetIterable(ByRef Elements As Variant) As String
    Dim Result As String
    Dim e As Variant
    For Each e In Elements
        Result = Result & ToString(e) & ", "
    Next

    If Strings.Len(Result) < 2 Then
        GetIterable = Result
    Else
        GetIterable = Strings.Left(Result, Strings.Len(Result) - 2)
    End If
End Function

Private Function GetString(ByVal Item As String) As String
    GetString = "'" & Item & "'"
End Function

Private Function GetDate(ByVal Item As String) As String
    GetDate = "#" & Item & "#"
End Function

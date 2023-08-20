Attribute VB_Name = "Strings"
'@Folder("VBA-Strings")
'@IgnoreModule ProcedureNotUsed
'Version(0.1)
Option Explicit

Public Enum Comparison
'vbUseCompareOption  -1 Performs a comparison by using the setting of the Option Compare statement.
'vbBinaryCompare     0  Performs a binary comparison.
'vbTextCompare       1  Performs a textual comparison.
'vbDatabaseCompare   2  Microsoft Access only. Performs a comparison based on information in your database.
'
'StringComparison
'CurrentCulture             0 Compare strings using culture-sensitive sort rules and the current culture.
'CurrentCultureIgnoreCase   1 Compare strings using culture-sensitive sort rules, the current culture, and ignoring the case of the strings being compared.
'InvariantCulture           2 Compare strings using culture-sensitive sort rules and the invariant culture.
'InvariantCultureIgnoreCase 3 Compare strings using culture-sensitive sort rules, the invariant culture, and ignoring the case of the strings being compared.
'Ordinal                    4 Compare strings using ordinal (binary) sort rules.
'OrdinalIgnoreCase          5 Compare strings using ordinal (binary) sort rules and ignoring the case of the strings being compared.
'
'params: 1) CompareType 2) IgnoreCase
    Default = 0
    DefaultIgnoreCase = 1
    Binary = 2
    BinaryIgnoreCase = 3
    Text = 4
    TextIngnoreCase = 5
    Database = 6
    DatabaseIgnoreCase = 7
End Enum



'TODO - easy way to convert between array and collection to make the _Any functions more flexible?


'https://stackoverflow.com/questions/4805475/assignment-of-objects-in-vb6/4805812#4805812
'Public Function Clone(ByRef baseString As String) As String
'    Clone = baseString
'End Function


Public Function Chars(ByVal stringToParse As String, ByVal index As Long)
    Char = Mid$(stringToParse, index, 1)
End Function


Public Sub CopyToCharArray(ByVal stringToCopy As String, ByRef Chars() As String)
    Dim idx As Long
    For idx = 1 To Len(stringToCopy)
        Chars(idx - 1) = Mid$(stringToCopy, idx, 1)
    Next idx
End Sub



'public void CopyTo (int sourceIndex, char[] destination, int destinationIndex, int count);
Public Sub CopyToCharArrayFrom(ByVal stringToCopy As String, ByVal sourceIndex As Long, ByRef Chars() As String, ByVal destinationIndex As Long, ByVal count As Long)
    Dim idx As Long
    Do
        Chars(destinationIndex + idx) = Mid$(stringToCopy, sourceIndex + idx + 1, 1)
        idx = idx + 1
    Loop While idx < count
End Sub




Public Function CountSubstring(ByVal stringToSearch As String, ByVal substringToFind As String, Optional compare As Comparison = 0) As Long
    Dim locn As Long: locn = IndexOf(stringToSearch, substringToFind, locn)
    
    Do While locn >= 0
        locn = IndexOf(stringToSearch, substringToFind, locn + 1)
        CountSubstring = CountSubstring + 1
    Loop
End Function





Public Function Create(ByVal Length As String, Optional ByVal defaultChar As String = " ") As String
    Create = Replace(Space(Length), " ", defaultChar)
End Function


Public Function EmptyString() As String
    EmptyString = vbNullString
End Function





Public Function Contains(ByVal searchstring As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    Contains = IndexOf(searchstring, stringToFind, , compare) >= 0
End Function



Public Function EndsWith(ByVal stringToSearch As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    If Comparison = 0 Then
        EndsWith = Right$(stringToSearch, Len(stringToFind)) = stringToFind
    Else
        EndsWith = IndexOfBetween(stringToSearch, stringToFind, Len(stringToSearch) - Len(stringToFind), Len(stringToFind), compare) >= 0
    End If
End Function



Public Function Equals(ByVal baseString As String, ByVal compareString As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    If ignoreCase Then
        baseString = LCase$(baseString)
        compareString = LCase$(compareString)
    End If
    Equals = baseString = compareString
End Function



Public Function GetTypeCode(ByVal item As Variant) As Long
    GetTypeCode = VarType(item)
End Function



'0-based index
Public Function IndexOf(ByVal searchstring As String, ByVal stringToFind As String, Optional ByVal startPos As Long = 0, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchstring = LCase$(searchstring)
        stringToFind = LCase$(stringToFind)
    End If
    
    If compare = Binary Or BinaryIgnoreCase Then
        IndexOf = InStr(startPos + 1, searchstring, stringToFind, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        IndexOf = InStr(startPos + 1, searchstring, stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        IndexOf = InStr(startPos + 1, searchstring, stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOf = InStr(startPos + 1, searchstring, stringToFind) - 1
    End If
End Function


Public Function IndexOfBetween(ByVal searchstring As String, ByVal stringToFind As String, ByVal startPos As Long, ByVal Length As Long, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchstring = LCase$(searchstring)
        stringToFind = LCase$(stringToFind)
    End If
    
    If compare = Binary Or BinaryIgnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchstring, startPos, Length), stringToFind, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchstring, startPos, Length), stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchstring, startPos, Length), stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOfBetween = InStr(1, Mid$(searchstring, startPos, Length), stringToFind) - 1
    End If
End Function


Public Function IndexOfAny(ByVal searchstring As String, ByRef stringsToFind() As String, Optional ByVal startPos As Long = 0, Optional ByVal Length As Long = 0) As Long
    If Length = 0 Then Length = Len(searchstring)
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        IndexOfAny = IndexOfBetween(searchstring, stringsToFind(idx), startPos + 1, Length)
        If IndexOfAny > 0 Then Exit Function
    Next idx
End Function



Public Function IsNull(ByVal searchstring As String) As Boolean
    IsNull = searchstring = vbNullString
End Function



Public Function IsNullOrWhiteSpace(ByRef searchstring As String) As Boolean
    'TODO - implement this
    'equivalent to: return String.IsNullOrEmpty(value) || value.Trim().Length == 0;
    IsNullOrWhiteSpace = False
End Function




Public Function LastIndexOf(ByVal searchstring As String, ByVal stringToFind As String, Optional ByVal startPos As Long = -2, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchstring = LCase$(searchstring)
        stringToFind = LCase$(stringToFind)
    End If
    
    If compare = Binary Or BinaryIgnoreCase Then
        LastIndexOf = InStrRev(searchstring, stringToFind, startPos + 1, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        LastIndexOf = InStrRev(searchstring, stringToFind, startPos + 1, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        LastIndexOf = InStrRev(searchstring, stringToFind, startPos + 1, vbDatabaseCompare) - 1
    Else
        LastIndexOf = InStrRev(searchstring, stringToFind, startPos + 1) - 1
    End If
End Function




Public Function LastIndexOfBetween(ByVal searchstring As String, ByVal stringToFind As String, ByVal startIndex As Long, ByVal Length As Long, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchstring = LCase$(searchstring)
        stringToFind = LCase$(stringToFind)
    End If
    
    If compare = Binary Or BinaryIgnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(searchstring, startIndex - Length, Length), stringToFind, -1, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(searchstring, startIndex - Length, Length), stringToFind, -1, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(searchstring, startIndex - Length, Length), stringToFind, -1, vbDatabaseCompare) - 1
    Else
        LastIndexOfBetween = InStrRev(Mid$(searchstring, startIndex - Length, Length), stringToFind, -1) - 1
    End If
    If LastIndexOfBetween > -1 Then
        LastIndexOfBetween = LastIndexOfBetween + startIndex - Length - 1
    End If
    
End Function


Public Function LastIndexOfAny(ByVal searchstring As String, ByRef stringsToFind() As String, Optional ByVal startPos As Long = 0, Optional ByVal Length As Long = 0) As Long
    If Length = 0 Then Length = Len(searchstring)
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        LastIndexOfAny = LastIndexOfBetween(searchstring, stringsToFind(idx), startPos + 1, Length)
        If LastIndexOfAny > 0 Then Exit Function
    Next idx
End Function



Public Function Left(ByVal stringToParse As String, ByVal count As Long) As String
    Left = VBA.Left$(stringToParse, count)
End Function




Public Function Length(ByVal stringToParse As String) As Long
    Length = Len(stringToParse)
End Function




Public Function Right(ByVal stringToParse As String, ByVal count As Long) As String
    Right = VBA.Right$(stringToParse, count)
End Function



Public Function StartsWith(ByVal stringToSearch As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    If Comparison = 0 Then
        StartsWith = Left$(stringToSearch, Len(stringToFind)) = stringToFind
    Else
        StartsWith = IndexOfBetween(stringToSearch, stringToFind, 0, Len(stringToFind), compare) >= 0
    End If
End Function






Public Sub test()
    Dim teststring As String
    teststring = "Public Sub ParseModule(moduleName As String)"
    Dim out As String
    out = Substring(teststring, 11, 11)
End Sub



Public Function SubstringBetween(ByVal stringToCut As String, ByVal firstString As String, ByVal secondString As String, Optional ByVal startIndex As Long = 0) As String
    Dim startPos As Integer: startPos = Strings.IndexOf(stringToCut, firstString, startIndex) + Strings.Length(firstString)
    Dim endPos As Integer: endPos = Strings.IndexOf(stringToCut, secondString, startPos)
    SubstringBetween = Strings.Substring(stringToCut, startPos, endPos - startPos)
End Function




Public Function Substring(ByVal stringToCut As String, ByVal atPosition As Long, Optional ByVal Length As Long = -1) As String
    If Length = -1 Then
        Substring = Mid$(stringToCut, atPosition + 1)
    Else
        Substring = Mid$(stringToCut, atPosition + 1, Length)
    End If
End Function


'Public Function ToCharArray(ByVal stringToCopy As String, Optional ByVal startIndex As Long = 0, Optional ByVal count As Long) As String()
'    Dim chars() As String
'    ReDim chars(0 To Len(stringToCopy))
'    Dim idx As Long
'    For idx = 1 To Len(stringToCopy)
'        chars(idx - 1) = Mid$(stringToCopy, idx, 1)
'    Next idx
'    CopyToArray = chars
'End Function





Public Function ToLower(ByVal stringToLowercase As String) As String
    ToLower = LCase$(stringToLowercase)
End Function




Public Function ToUpper(ByVal stringToUppercase As String) As String
    ToUpper = UCase$(stringToUppercase)
End Function


Public Function Trim(ByVal stringToTrim As String) As String
    Trim = VBA.Trim$(stringToTrim)
End Function


Public Function TrimEnd(ByVal stringToTrim As String) As String
    TrimEnd = VBA.RTrim$(stringToTrim)
End Function

Public Function TrimLeft(ByVal stringToTrim As String) As String
    TrimLeft = VBA.LTrim$(stringToTrim)
End Function

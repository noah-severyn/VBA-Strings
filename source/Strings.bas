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




'AscW Fix
'The AscW function in the built-in VBA.Strings module has a problem where it returns the correct bit pattern for an unsigned 16-bit integer which is
'incorrect in VBA because VBA uses signed 16-bit integer. Thus, after reaching 32767 AscW will start returning negative numbers. To work around this
'issue use one of the functions below.
'@Ignore UseMeaningfulName
Public Function AscW2(ByVal Char As String) As Long
    AscW2 = AscW(Char) And &HFFFF&
End Function


Public Function Append(ByVal baseString As String, ParamArray args() As Variant) As String
    Append = baseString
    Dim argIdx As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To UBound(args)
        Append = Append & args(argIdx)
    Next argIdx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.Append", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function

Public Function Append2(ByVal baseString As String, ByVal args As Collection) As String
    Append = baseString
    Dim argIdx As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To args.count
        Append = Append & args(argIdx)
    Next argIdx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.Append", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function


Public Function AppendLine(ByVal baseString As String, ParamArray args() As Variant) As String
    Append = baseString
    Dim argIdx As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To UBound(args)
        Append = Append & Chr$(10) & args(argIdx)
    Next argIdx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.Append", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function

Public Function AppendLine2(ByVal baseString As String, ByVal args As Collection) As String
    Append = baseString
    Dim argIdx As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To args.count
        Append = Append & Chr$(10) & args(argIdx)
    Next argIdx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.Append", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function


'https://stackoverflow.com/questions/4805475/assignment-of-objects-in-vb6/4805812#4805812
'Public Function Clone(ByRef baseString As String) As String
'    Clone = baseString
'End Function


Public Function Char(ByVal baseString As String, ByVal index As Long) As String
    Char = Mid$(baseString, index, 1)
End Function



Public Function Clean(ByVal baseString As String, Optional ByVal nonPrintable As Boolean = True, Optional ByVal newLines As Boolean = True, Optional ByVal nonBreaking As Boolean = True, Optional ByVal trimString As Boolean = True, Optional ByVal newLineReplacement As String = " ") As String
    Clean = baseString
    If nonPrintable Then Clean = strings.RemoveNonPrintableChars(Clean)
    If newLines Then Clean = strings.ReplaceNewLineChars(Clean, newLineReplacement)
    If nonBreaking Then Clean = strings.ReplaceNonBreakingSpaces(Clean)
    If trimString Then Clean = strings.Trim(Clean)
End Function


Public Function Clear(ByVal baseString As String)
    Clear = vbNullString
End Function




'nullstring, null
'Returns the first non null parameter
Public Function Coalesce(ParamArray params() As Variant) As String
    Dim idx As Long
    Dim currentParam As Variant
    For idx = 0 To UBound(params)
        currentParam = params(idx)
        Coalesce = currentParam
        If Not IsNull(currentParam) And currentParam <> vbNullString Then
            Exit Function
        End If
    Next idx
End Function


Public Function Concat(ByVal delimiter As String, ParamArray args() As Variant) As String
    Dim idx As Long
    Dim idxInner As Long
    For idx = LBound(args) To UBound(args)
        If VBA.IsArray(args(idx)) Then
            For idxInner = LBound(args(idx)) To UBound(args(idx))
                Concat = Concat & CStr(args(idx)(idxInner))
                If idxInner < UBound(args(idx)) - 1 Then Concat = Concat & delimiter
            Next idxInner
        Else
            Concat = Concat & CStr(args(idx))
            If idx < UBound(args) Then Concat = Concat & delimiter
        End If
    Next idx
End Function




Public Function Contains(ByVal searchString As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    Contains = IndexOf(searchString, stringToFind, , compare) >= 0
End Function


Public Function ContainsAfter(ByVal searchString As String, ByVal stringToFind As String, Optional ByVal startIndex As Long, Optional ByVal compare As Comparison = 0) As Boolean
    ContainsAfter = IndexOf(searchString, stringToFind, startIndex, compare) >= 0
End Function



Public Function ContainsBefore(ByVal searchString As String, ByVal stringToFind As String, Optional ByVal endIndex As Long, Optional ByVal compare As Comparison = 0) As Boolean
    Dim newString As String: newString = Left(searchString, endIndex)
    ContainsBefore = IndexOf(newString, stringToFind, , compare) >= 0
End Function




Public Function ContainsAny(ByVal searchString As String, ByVal stringsToFind As Variant, Optional ByVal startIndex As Long, Optional ByVal compare As Comparison = 0) As Boolean
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        ContainsAny = IndexOf(searchString, stringsToFind(idx), startIndex, compare) >= 0
        If ContainsAny Then Exit Function
    Next idx
End Function





'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub CopyToCharArray(ByVal stringToCopy As String, ByRef Chars() As String)
    Dim idx As Long
    For idx = 1 To Len(stringToCopy)
        Chars(idx - 1) = Mid$(stringToCopy, idx, 1)
    Next idx
End Sub



'public void CopyTo (int sourceIndex, char[] destination, int destinationIndex, int count);
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub CopyToCharArrayFrom(ByVal stringToCopy As String, ByVal sourceIndex As Long, ByRef Chars() As String, ByVal destinationIndex As Long, ByVal count As Long)
    Dim idx As Long
    Do
        Chars(destinationIndex + idx) = Mid$(stringToCopy, sourceIndex + idx + 1, 1)
        idx = idx + 1
    Loop While idx < count
End Sub




Public Function CountSubstring(ByVal stringToSearch As String, ByVal substringToFind As String, Optional ByVal compare As Comparison = 0) As Long
    Dim locn As Long: locn = IndexOf(stringToSearch, substringToFind, locn)

    Do While locn >= 0
        locn = IndexOf(stringToSearch, substringToFind, locn + 1, compare)
        CountSubstring = CountSubstring + 1
    Loop
End Function





Public Function Create(ByVal Length As Long, Optional ByVal defaultChar As String = " ") As String
    Create = Replace(space(Length), " ", defaultChar)
End Function


Public Function EmptyString() As String
    EmptyString = vbNullString
End Function




Public Function EndsWith(ByVal stringToSearch As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    If Comparison = 0 Then
        EndsWith = Right$(stringToSearch, Len(stringToFind)) = stringToFind
    Else
        EndsWith = IndexOfBetween(stringToSearch, stringToFind, Len(stringToSearch) - Len(stringToFind), Len(stringToFind), compare) >= 0
    End If
End Function


Public Function EndsWithAny(ByVal stringToSearch As String, ByVal stringsToFind As Variant, Optional ByVal compare As Comparison = 0) As Boolean
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        EndsWithAny = EndsWith(stringToSearch, stringsToFind(idx), compare)
        If EndsWithAny Then Exit Function
    Next idx
End Function







Public Function Equals(ByVal baseString As String, ByVal compare As Comparison, ByVal compareString As String) As Boolean
    If compare Mod 2 = 1 Then
        baseString = LCase$(baseString)
        compareString = LCase$(compareString)
    End If
    
    Matches = baseString = compareString
End Function

Public Function EqualsAny(ByVal baseString As String, ByVal compare As Comparison, ParamArray compareStrings() As Variant) As Boolean
    Dim idx As Long
    For idx = LBound(compareStrings) To UBound(compareStrings)
        MatchesAny = Matches(baseString, compare, compareStrings(idx))
        If MatchesAny = True Then Exit Function
    Next idx
End Function



Public Function GetTypeCode(ByVal item As Variant) As Long
    GetTypeCode = VarType(item)
End Function



Public Function IndexOf(ByVal searchString As String, ByVal stringToFind As String, Optional ByVal startPos As Long = 0, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchString = LCase$(searchString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or BinaryIgnoreCase Then
        IndexOf = InStr(startPos + 1, searchString, stringToFind, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        IndexOf = InStr(startPos + 1, searchString, stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        IndexOf = InStr(startPos + 1, searchString, stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOf = InStr(startPos + 1, searchString, stringToFind) - 1
    End If
End Function


Public Function IndexOfBetween(ByVal searchString As String, ByVal stringToFind As String, ByVal startPos As Long, ByVal Length As Long, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchString = LCase$(searchString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or BinaryIgnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, Length), stringToFind, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, Length), stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, Length), stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, Length), stringToFind) - 1
    End If
End Function


Public Function IndexOfAny(ByVal searchString As String, ByRef stringsToFind() As String, Optional ByVal startPos As Long = 0, Optional ByVal Length As Long = 0) As Long
    If Length = 0 Then Length = Len(searchString)
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        IndexOfAny = IndexOfBetween(searchString, stringsToFind(idx), startPos + 1, Length)
        If IndexOfAny > 0 Then Exit Function
    Next idx
End Function



Public Function Insert(ByVal baseString As String, ByVal startIndex As Long, ByVal stringToInsert As String) As String
    Insert = Left$(baseString, startIndex) & stringToInsert & Right$(baseString, Len(baseString) - startIndex)
End Function



Public Function Interpolate(ByVal baseString As String, ParamArray args() As Variant) As String
    Dim argCount As Long: argCount = UBound(args) - LBound(args) + 1
    If baseString = vbNullString Then
        Err.Raise 5, "Strings.Interpolate", "Base string cannot be null."
    ElseIf (argCount) Mod 2 <> 0 Then
        Err.Raise 5, "Strings.Interpolate", "Invalid number of parameters. The interpolation parameters must be provided in pairs."
    End If
    
    Dim argIdx As Long
    Interpolate = baseString
    On Error GoTo InterpolateInvalidArg
    For argIdx = 0 To argCount - 2 Step 2 '
        Interpolate = strings.Replace(Interpolate, "{" & CStr(args(argIdx)) & "}", CStr(args(argIdx + 1)))
    Next argIdx
    On Error GoTo 0
    Exit Function
    
InterpolateInvalidArg:
    Err.Raise 5, "Strings.Interpolate", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function


Public Function Interpolate2(ByVal baseString As String, ByVal args As Dictionary)
    If baseString = vbNullString Then
        Err.Raise 5, "Strings.Interpolate", "Base string cannot be null."
    End If
    
    Dim argIdx As Long
    Interpolate2 = baseString
    Dim key As Variant
    On Error GoTo Interpolate2InvalidArg
    For Each key In args.Keys()
        Interpolate2 = strings.Replace(Interpolate2, "{" & CStr(key) & "}", CStr(args(key)))
    Next key
    On Error GoTo 0
    Exit Function
    
Interpolate2InvalidArg:
    Err.Raise 5, "Strings.Interpolate", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function




Public Function IsNull(ByVal searchString As String) As Boolean
    IsNull = searchString = vbNullString
End Function



Public Function IsNullOrWhiteSpace(ByVal searchString As String) As Boolean
    'TODO - .IsNullOrWhiteSpace implement this
    'equivalent to: return String.IsNullOrEmpty(value) || value.Trim().Length == 0;
    IsNullOrWhiteSpace = False
End Function



Public Function Join(ByVal delimiter As String, ByRef stringsToJoin() As String) As String
    Join = VBA.Join(stringsToJoin, delimiter)
End Function


Public Function JoinBetween(ByVal delimiter As String, ByRef stringsToJoin() As String, ByVal startIndex As Long, ByVal count As Long) As String
    If startIndex < 0 Then
        Err.Raise 9, "Strings.JoinBetween", "Invalid startIndex"
    ElseIf count < 0 Then
        Err.Raise 5, "Strings.JoinBetween", "Invalid count"
    ElseIf count = 0 Then
        Exit Function
    End If
    
    Dim subset() As String
    ReDim subset(0 To count - 1) As String
    Dim idx As Long
    For idx = 0 To count - 1
        subset(idx) = stringsToJoin(idx + startIndex)
    Next idx
    
    JoinBetween = strings.Join(delimiter, subset)
End Function





Public Function LastIndexOf(ByVal searchString As String, ByVal stringToFind As String, Optional ByVal startPos As Long = -2, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchString = LCase$(searchString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or BinaryIgnoreCase Then
        LastIndexOf = InStrRev(searchString, stringToFind, startPos + 1, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        LastIndexOf = InStrRev(searchString, stringToFind, startPos + 1, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        LastIndexOf = InStrRev(searchString, stringToFind, startPos + 1, vbDatabaseCompare) - 1
    Else
        LastIndexOf = InStrRev(searchString, stringToFind, startPos + 1) - 1
    End If
End Function




Public Function LastIndexOfBetween(ByVal searchString As String, ByVal stringToFind As String, ByVal startIndex As Long, ByVal Length As Long, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchString = LCase$(searchString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or BinaryIgnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(searchString, startIndex - Length, Length), stringToFind, -1, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(searchString, startIndex - Length, Length), stringToFind, -1, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(searchString, startIndex - Length, Length), stringToFind, -1, vbDatabaseCompare) - 1
    Else
        LastIndexOfBetween = InStrRev(Mid$(searchString, startIndex - Length, Length), stringToFind, -1) - 1
    End If
    If LastIndexOfBetween > -1 Then
        LastIndexOfBetween = LastIndexOfBetween + startIndex - Length - 1
    End If

End Function


Public Function LastIndexOfAny(ByVal searchString As String, ByRef stringsToFind() As String, Optional ByVal startPos As Long = 0, Optional ByVal Length As Long = 0) As Long
    If Length = 0 Then Length = Len(searchString)
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        LastIndexOfAny = LastIndexOfBetween(searchString, stringsToFind(idx), startPos + 1, Length)
        If LastIndexOfAny > 0 Then Exit Function
    Next idx
End Function



Public Function Left(ByVal stringToParse As String, ByVal count As Long) As String
    Left = VBA.Left$(stringToParse, count)
End Function




Public Function Length(ByVal stringToParse As String) As Long
    Length = Len(stringToParse)
End Function



'/**
' * Levenshtein is the distance between two sequences of words.
' *
' * @author Robert Todar <robert@robertodar.com>
' * @see <https://www.cuelogic.com/blog/the-levenshtein-algorithm>
' * @example LevenshteinDistance("Test", "Tester") ->  2
' */
Public Function LevenshteinDistance(ByVal firstString As String, ByVal secondString As String) As Double
    Dim firstLength As Long: firstLength = Len(firstString)
    Dim secondLength As Long: secondLength = Len(secondString)

    ' Prepare distance array matrix with the proper indexes
    Dim distance() As Long
    ReDim distance(firstLength, secondLength)

    Dim index As Long
    For index = 0 To firstLength
        distance(index, 0) = index
    Next

    Dim innerIndex As Long
    For innerIndex = 0 To secondLength
        distance(0, innerIndex) = innerIndex
    Next

    ' Outer loop is for the first string
    For index = 1 To firstLength
        ' Inner loop is for the second string
        For innerIndex = 1 To secondLength
            ' Character matches exactly
            If Mid$(firstString, index, 1) = Mid$(secondString, innerIndex, 1) Then
                distance(index, innerIndex) = distance(index - 1, innerIndex - 1)

            ' Character is off, offset the matrix by the appropriate number.
            Else
                Dim min1 As Long: min1 = distance(index - 1, innerIndex) + 1
                Dim min2 As Long: min2 = distance(index, innerIndex - 1) + 1

                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = distance(index - 1, innerIndex - 1) + 1

                If min2 < min1 Then
                    min1 = min2
                End If
                distance(index, innerIndex) = min1
            End If
        Next
    Next

    ' Levenshtein is the last index of the array.
    LevenshteinDistance = distance(firstLength, secondLength)
End Function



'/**
' * This returns a percentage of how similar two strings are using the levenshtein formula.
' *
' * @author Robert Todar <robert@robertodar.com>
' * @example StringSimilarity("Test", "Tester") ->  66.6666666666667
' */
Public Function MeasureSimilarity(ByVal firstString As String, ByVal secondString As String) As Double
    ' Levenshtein is the distance between two sequences
    Dim levenshtein As Double
    levenshtein = LevenshteinDistance(firstString, secondString)

    ' Convert levenshtein into a percentage(0 to 100)
    MeasureSimilarity = (1 - (levenshtein / Application.Max(Len(firstString), Len(secondString)))) * 100
End Function



Public Function Overwrite(ByVal baseString As String, ByVal startIndex As Long, ByVal replacementText As String)
    'Validate Input
    If startIndex < 0 Then
        Err.Raise 9, "Strings.Overwrite", "Start index cannot be less than zero."
    ElseIf startIndex > Len(baseString) < 0 Then
        Err.Raise 5, "Strings.Overwrite", "Start index cannot be longer than the length of the base string."
    End If
    
    Overwrite = Left(baseString, startIndex) & replacementText & Right(baseString, Len(baseString) - Len(replacementText))
End Function




Public Function PadLeft(ByVal baseString As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String = " ") As String
    If Len(paddingChar) > 1 Then
        Err.Raise 9, "Strings.PadLeft", "Padding character can only be one character in length"
    ElseIf paddingChar = vbNullString Then
        Err.Raise 9, "Strings.PadLeft", "Padding character cannot be null"
    End If
    
    PadLeft = VBA.String$(totalWidth - Len(baseString), paddingChar) & baseString
End Function




Public Function PadRight(ByVal baseString As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String = " ") As String
    If Len(paddingChar) > 1 Then
        Err.Raise 9, "Strings.PadRight", "Padding character can only be one character in length"
    ElseIf paddingChar = vbNullString Then
        Err.Raise 9, "Strings.PadRight", "Padding character cannot be null"
    End If
    
    PadRight = baseString & VBA.String$(totalWidth - Len(baseString), paddingChar)
End Function




Public Function Remove(ByVal baseString As String, ByVal startIndex As Long, Optional ByVal count = 0) As String
    If startIndex < 0 Then
        Err.Raise 9, "Strings.Remove", "Start index must be greater than 0"
    ElseIf count < 0 Then
        Err.Raise 9, "Strings.Remove", "Count must be greater than 0"
    End If
    
    If count = 0 Then
        Remove = strings.Left(baseString, count)
    Else
        Remove = strings.Left(baseString, Len(baseString) - startIndex)
    End If
End Function



Public Function RemoveFromEndWhile(ByVal baseString As String, ByVal stringToRemove As String)
    If stringToRemove = vbNullString Or Len(stringToRemove) = 0 Then
        Err.Raise 9, "Strings.RemoveFromEndWhile", "String to remove must have length greater than 0 and cannot be null"
    End If
    
    Dim charCount As Long: charCount = Len(stringToRemove)
    RemoveFromEndWhile = baseString
    
    Do While strings.Right(RemoveFromEndWhile, charCount) = stringToRemove
        RemoveFromEndWhile = strings.Left(baseString, Len(RemoveFromEndWhile) - charCount)
    Loop
End Function



Public Function RemoveNonPrintableChars(ByVal baseString As String) As String
    'Does not remove new line characters
    Dim idx As Long
    Dim c As Long
    Dim currentCharCode As Long

    RemoveNonPrintableChars = String$(Len(baseString), Chr$(0))
    For idx = 1 To Len(Text)
        currentCharCode = AscW2(Mid$(baseString, idx, 1))
        If currentCharCode > 31 Or currentCharCode = 13 Or currentCharCode = 10 Then
            c = c + 1
            Mid$(RemoveNonPrintableChars, c, 1) = Mid$(baseString, idx, 1)
        End If
    Next idx
    RemoveNonPrintableChars = Left$(RemoveNonPrintableChars, c)
End Function



Public Function Replace(ByVal baseString As String, ByVal oldString As String, ByVal newString As String, Optional ByVal compare As Comparison = Comparison.Default)
    If oldString = vbNullString Then
        Err.Raise 9, "Strings.Replace", "String to replace cannot be null"
    End If
    
    If compare Mod 2 = 1 Then 'TODO - .Replace case comparison probably isn't working
        baseString = LCase$(baseString)
        'stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or BinaryIgnoreCase Then
        Replace = VBA.Replace(baseString, oldString, newString, , , vbBinaryCompare)
    ElseIf compare = Text Or TextIngnoreCase Then
        Replace = VBA.Replace(baseString, oldString, newString, , , vbTextCompare)
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        Replace = VBA.Replace(baseString, oldString, newString, , , vbDatabaseCompare)
    Else
        Replace = VBA.Replace(baseString, oldString, newString)
    End If
    
End Function






Public Function Right(ByVal baseString As String, ByVal count As Long) As String
    Right = VBA.Right$(baseString, count)
End Function



Public Function StartsWith(ByVal stringToSearch As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    If Comparison = 0 Then
        StartsWith = Left$(stringToSearch, Len(stringToFind)) = stringToFind
    Else
        StartsWith = IndexOfBetween(stringToSearch, stringToFind, 0, Len(stringToFind), compare) >= 0
    End If
End Function


Public Function StartsWithAny(ByVal stringToSearch As String, ByVal stringsToFind As Variant, Optional ByVal compare As Comparison = 0) As Boolean
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        StartsWithAny = StartsWith(stringToSearch, stringsToFind(idx), compare)
        If StartsWithAny Then Exit Function
    Next idx
End Function



'Public Sub test()
'    Dim teststring As String
'    teststring = "Public Sub ParseModule(moduleName As String)"
'    Dim out As String
'    out = Substring(teststring, 11, 11)
'End Sub



Public Function SubstringBetween(ByVal stringToCut As String, ByVal firstString As String, ByVal secondString As String, Optional ByVal startIndex As Long = 0) As String
    Dim startPos As Long: startPos = strings.IndexOf(stringToCut, firstString, startIndex) + strings.Length(firstString)
    Dim endPos As Long: endPos = strings.IndexOf(stringToCut, secondString, startPos)
    SubstringBetween = strings.Substring(stringToCut, startPos, endPos - startPos)
End Function




Public Function Substring(ByVal stringToCut As String, ByVal atPosition As Long, Optional ByVal Length As Long = -1) As String
    If Length = -1 Then
        Substring = Mid$(stringToCut, atPosition + 1)
    Else
        Substring = Mid$(stringToCut, atPosition + 1, Length)
    End If
End Function


Public Function Reverse(ByVal baseString As String)
    Reverse = VBA.StrReverse$(baseString)
End Function




Public Function ReplaceNonBreakingSpaces(ByVal baseString As String) As String
    ReplaceNonBreakingSpaces = Replace(baseString, Chr$(160), " ")
End Function




Public Function ReplaceNewLineChars(ByVal baseString As String, Optional ByVal replacement As String = " ") As String
    ReplaceNewLineChars = VBA.Replace$(baseString, vbCrLf, replacement)
    ReplaceNewLineChars = VBA.Replace$(ReplaceNewLineChars, vbCr, replacement)
    ReplaceNewLineChars = VBA.Replace$(ReplaceNewLineChars, vbLf, replacement)
End Function


Public Function Split(ByVal baseString As String, ParamArray delimiters() As Variant) As String()
    Dim delims() As String
    strings.ToStringArray delims, delimiters 'TODO - .Split bug here
    
    Dim result() As String
    ReDim result(0 To Len(baseString * 2)) 'worse case scenario?
    
    Dim pos As Long
    Dim startPos As Long
    Dim outIdx As Long
    Dim idx As Long
    For idx = 0 To UBound(result)
        startPos = pos
        pos = strings.IndexOfAny(searchString, delims, pos)
        result(idx) = strings.Substring(baseString, startPos, pos - startPos)
    Next idx
End Function



Public Function ToCharArray(ByVal stringToSplit As String, Optional ByVal sourceIndex As Long, Optional ByVal count As Long) As String()
    Dim Chars() As String
    ReDim Chars(0 To Len(stringToSplit)) As String
    Dim idx As Long
    Do
        Chars(sourceIndex + idx) = Mid$(stringToSplit, sourceIndex + idx + 1, 1)
        idx = idx + 1
    Loop While idx < count
    ToCharArray = Chars
End Function


Public Function ToStringArray(ByRef outputArray() As String, ByRef inputArray() As Variant) As String()
    ReDim outputArray(LBound(inputArray) To UBound(inputArray))
    Dim idx
    On Error GoTo ErrorCannotConvertToString:
    For idx = 0 To UBound(outputArray)
        outputArray(idx) = CStr(inputArray(idx))
        idx2 = idx2 + 1
    Next idx
    On Error GoTo 0
    Exit Function
    
ErrorCannotConvertToString:
    Err.Raise 13, "Strings.ToStringArray", "Cannot convert element at "
    On Error GoTo 0
End Function




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


'/**
' * Create a max length of string and return it with extension.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @example Truncate("This is a long sentence", 10)  -> "This is..."
' */
Public Function Truncate(ByRef source As String, ByVal maxLength As Long, Optional ByVal extension = "...") As String
    If Len(source) <= maxLength Then
        Truncate = source
        Exit Function
    End If

    source = Left(source, maxLength - Len(extention)) & extention
    Truncate = source
End Function

'from https://stackoverflow.com/a/218199/10802255
Public Function URLEncode(stringToEncode As String, Optional spaceAsPlus As Boolean = False) As String
    Dim stringLen As Long: stringLen = Len(stringToEncode)

    If stringLen > 0 Then
        ReDim result(stringLen) As String
        Dim i As Long
        Dim charCode As Integer
        Dim Char As String
        Dim space As String

        If spaceAsPlus Then space = "+" Else space = "%20"

        For i = 1 To stringLen
            Char = Mid$(stringToEncode, i, 1)
            charCode = Asc(Char)
            Select Case charCode
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                result(i) = Char
            Case 32
                result(i) = space
            Case 0 To 15
                result(i) = "%0" & Hex$(charCode)
            Case Else
                result(i) = "%" & Hex$(charCode)
            End Select
        Next i
        URLEncode = Join(result, "")
    End If
End Function

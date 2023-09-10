Attribute VB_Name = "Strings"
'@Folder("VBA-Strings")
'@IgnoreModule ProcedureNotUsed
''===================================================================================================================================================
'' VBA-Strings
'' Version 0.2
''-------------------------------------------------
'' https://github.com/noah-severyn/VBA-Strings
''-------------------------------------------------
''
'' Copyright (c) 2023 Noah Severyn
''
'' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
'' to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
'' and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
''
'' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
''
'' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
'' IN THE SOFTWARE.
''===================================================================================================================================================
''===================================================================================================================================================
'' Description:
''    * This module is an collection of string manipulation functions based on the ones provided by the .NET built-in String class.
''    * For usage instructions, including for when to use regular strings over the StringBuilder class, refer to the documentation at:
''      https://github.com/noah-severyn/VBA-Strings
''===================================================================================================================================================
Option Explicit


Public Enum Comparison
    Default = 0
    DefaultIgnoreCase = 1
    Binary = 2
    BinaryIgnoreCase = 3
    text = 4
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
    Dim argIdxInner As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To UBound(args)
        If IsArray(args(argIdx)) Then
            For argIdxInner = 0 To UBound(args(argIdx))
                Append = Append & CStr(args(argIdx)(argIdxInner))
            Next argIdxInner
        Else
            Append = Append & CStr(args(argIdx))
        End If
    Next argIdx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.Append", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function

Public Function Append2(ByVal baseString As String, ByVal args As Collection) As String
    Append2 = baseString
    Dim argIdx As Long
    Dim argIdxInner As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To args.count
        If IsArray(args.item(argIdx)) Then
            For argIdxInner = 0 To UBound(args.item(argIdx))
                Append2 = Append2 & CStr(args.item(argIdx)(argIdxInner))
            Next argIdxInner
        Else
            Append2 = Append2 & CStr(args.item(argIdx))
        End If
    Next argIdx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.Append2", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function


Public Function AppendLine(ByVal baseString As String, ParamArray args() As Variant) As String
    AppendLine = baseString
    Dim argIdx As Long
    Dim argIdxInner As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To args.count
        If IsArray(args(argIdx)) Then
            For argIdxInner = 0 To UBound(args(argIdx))
                AppendLine = AppendLine & vbNewLine & CStr(args(argIdx)(argIdxInner))
            Next argIdxInner
        Else
            AppendLine = AppendLine & vbNewLine & CStr(args(argIdx))
        End If
    Next argIdx
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.AppendLine", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function

Public Function AppendLine2(ByVal baseString As String, ByVal args As Collection) As String
    AppendLine2 = baseString
    Dim argIdx As Long
    Dim argIdxInner As Long
    On Error GoTo AppendInvalidArg
    For argIdx = 0 To args.count
        If IsArray(args.item(argIdx)) Then
            For argIdxInner = 0 To UBound(args.item(argIdx))
                AppendLine2 = AppendLine2 & vbNewLine & CStr(args.item(argIdx)(argIdxInner))
            Next argIdxInner
        Else
            AppendLine2 = AppendLine2 & vbNewLine & CStr(args.item(argIdx))
        End If
    Next argIdx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.AppendLine2", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function


'https://stackoverflow.com/questions/4805475/assignment-of-objects-in-vb6/4805812#4805812
'Public Function Clone(ByRef baseString As String) As String
'    Clone = baseString
'End Function


Public Function Char(ByVal baseString As String, ByVal index As Long) As String
    If index < 0 Then
        Err.Raise 9, "Strings.Char", "Index must be greater than zero."
    ElseIf (index) > Len(baseString) Then
        index = Len(baseString) - 1
    End If

    Char = Mid$(baseString, index + 1, 1)
End Function



Public Function Clean(ByVal baseString As String, Optional ByVal nonPrintable As Boolean = True, Optional ByVal newLines As Boolean = True, Optional ByVal nonBreaking As Boolean = True, Optional ByVal trimString As Boolean = True, Optional ByVal newLineReplacement As String = " ") As String
    Clean = baseString
    If nonPrintable Then Clean = Strings.RemoveNonPrintableChars(Clean)
    If newLines Then Clean = Strings.ReplaceNewLineChars(Clean, newLineReplacement)
    If nonBreaking Then Clean = Strings.ReplaceNonBreakingSpaces(Clean)
    If trimString Then Clean = Strings.Trim(Clean)
End Function


Public Function Clear(ByVal baseString As String) As String
    Clear = vbNullString
End Function




'nullstring, null
'Returns the first non null parameter
Public Function Coalesce(ParamArray args() As Variant) As String
    Dim idx As Long
    Dim currentParam As Variant
    For idx = 0 To UBound(args)
        currentParam = args(idx)
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




Public Function Contains(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    Contains = IndexOf(baseString, stringToFind, , compare) >= 0
End Function


Public Function ContainsAfter(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal startIndex As Long, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    If startIndex < 0 Then
        Err.Raise 9, "Strings.ContainsAfter", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.ContainsAfter", "Start index must be less than the base string length."
    End If
    
    ContainsAfter = IndexOf(baseString, stringToFind, startIndex, compare) >= 0
End Function



Public Function ContainsBefore(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal endIndex As Long, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    If endIndex < 0 Then
        Err.Raise 9, "Strings.ContainsBefore", "End index must be greater than zero."
    End If
    
    Dim newString As String: newString = Left(baseString, endIndex)
    ContainsBefore = IndexOf(newString, stringToFind, , compare) >= 0
End Function




Public Function ContainsAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal startIndex As Long, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    If startIndex < 0 Then
        Err.Raise 9, "Strings.ContainsAfter", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.ContainsAfter", "Start index must be less than the base string length."
    End If
    
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        ContainsAny = IndexOf(baseString, stringsToFind(idx), startIndex, compare) >= 0
        If ContainsAny Then Exit Function
    Next idx
End Function





'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub CopyToCharArray(ByVal baseString As String, ByRef Chars() As String)
    Dim idx As Long
    For idx = 1 To Len(baseString)
        Chars(idx - 1) = Mid$(baseString, idx, 1)
    Next idx
End Sub



'public void CopyTo (int sourceIndex, char[] destination, int destinationIndex, int count);
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub CopyToCharArrayFrom(ByVal baseString As String, ByVal sourceIndex As Long, ByRef charArray() As String, ByVal destinationIndex As Long, ByVal count As Long)
    Dim idx As Long
    Do
        charArray(destinationIndex + idx) = Mid$(baseString, sourceIndex + idx + 1, 1)
        idx = idx + 1
    Loop While idx < count
End Sub




Public Function CountSubstring(ByVal baseString As String, ByVal substringToFind As String, Optional ByVal compare As Comparison = Comparison.Default) As Long
    Dim locn As Long: locn = IndexOf(baseString, substringToFind, locn)

    Do While locn >= 0
        locn = IndexOf(baseString, substringToFind, locn + 1, compare)
        CountSubstring = CountSubstring + 1
    Loop
End Function





Public Function Create(ByVal count As Long, Optional ByVal defaultChar As String = " ") As String
    Create = Replace(space(count), " ", defaultChar)
End Function


Public Function EmptyString() As String
    EmptyString = vbNullString
End Function




Public Function EndsWith(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    If Comparison = Comparison.Default Then
        EndsWith = Right$(baseString, Len(stringToFind)) = stringToFind
    Else
        EndsWith = IndexOfBetween(baseString, stringToFind, Len(baseString) - Len(stringToFind), Len(stringToFind), compare) >= 0
    End If
End Function


Public Function EndsWithAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        EndsWithAny = EndsWith(baseString, stringsToFind(idx), compare)
        If EndsWithAny Then Exit Function
    Next idx
End Function







Public Function Equals(ByVal baseString As String, ByVal compareString As String, ByVal compare As Comparison) As Boolean
    If compare Mod 2 = 1 Then
        baseString = LCase$(baseString)
        compareString = LCase$(compareString)
    End If
    
    Equals = baseString = compareString
End Function

Public Function EqualsAny(ByVal baseString As String, ByVal compare As Comparison, ParamArray compareStrings() As Variant) As Boolean
    Dim idx As Long
    For idx = LBound(compareStrings) To UBound(compareStrings)
        EqualsAny = Equals(baseString, compare, compareStrings(idx))
        If EqualsAny = True Then Exit Function
    Next idx
End Function



Public Function GetTypeCode(ByVal item As Variant) As Long
    GetTypeCode = VarType(item)
End Function



Public Function IndexOf(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal startPos As Long = 0, Optional ByVal compare As Comparison = Comparison.Default) As Long
    If startPos < 0 Then
        Err.Raise 9, "Strings.IndexOf", "Start position must be greater than zero."
    ElseIf startPos > Len(baseString) Then
        Err.Raise 9, "Strings.IndexOf", "Start position must be less than the base string length."
    End If
    
    If compare Mod 2 = 1 Then
        baseString = LCase$(baseString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or compare = BinaryIgnoreCase Then
        IndexOf = InStr(startPos + 1, baseString, stringToFind, vbBinaryCompare) - 1
    ElseIf compare = text Or compare = TextIngnoreCase Then
        IndexOf = InStr(startPos + 1, baseString, stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or compare = DatabaseIgnoreCase Then
        IndexOf = InStr(startPos + 1, baseString, stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOf = InStr(startPos + 1, baseString, stringToFind) - 1
    End If
End Function


Public Function IndexOfBetween(ByVal baseString As String, ByVal stringToFind As String, ByVal startPos As Long, ByVal count As Long, Optional ByVal compare As Comparison = Comparison.Default) As Long
    If startPos < 0 Then
        Err.Raise 9, "Strings.IndexOfBetween", "Start position must be greater than zero."
    ElseIf startPos > Len(baseString) Then
        Err.Raise 9, "Strings.IndexOfBetween", "Start position must be less than the base string length."
    ElseIf startPos + count > Len(baseString) Then
        count = Len(baseString) - startPos
    End If
    
    If compare Mod 2 = 1 Then
        baseString = LCase$(baseString)
        stringToFind = LCase$(stringToFind)
    End If
    startPos = startPos + 1 'Adjustment for the 1-based strings in VBA

    If compare = Binary Or compare = BinaryIgnoreCase Then
        IndexOfBetween = InStr(startPos, Mid$(baseString, startPos, count), stringToFind, vbBinaryCompare) - 1
    ElseIf compare = text Or compare = TextIngnoreCase Then
        IndexOfBetween = InStr(startPos, Mid$(baseString, startPos, count), stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or compare = DatabaseIgnoreCase Then
        IndexOfBetween = InStr(startPos, Mid$(baseString, startPos, count), stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOfBetween = InStr(startPos, Mid$(baseString, startPos, count), stringToFind) - 1
    End If
End Function


Public Function IndexOfAny(ByVal baseString As String, ByRef stringsToFind() As String, Optional ByVal startPos As Long = 0, Optional ByVal count As Long = 0) As Long
    If startPos < 0 Then
        Err.Raise 9, "Strings.IndexOfAny", "Start position must be greater than zero."
    ElseIf startPos > Len(baseString) Then
        Err.Raise 9, "Strings.IndexOfAny", "Start position must be less than the base string length."
    ElseIf startPos + count > Len(baseString) Then
        count = Len(baseString) - startPos
    End If
    
    If count = 0 Then count = Len(baseString)
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        IndexOfAny = IndexOfBetween(baseString, stringsToFind(idx), startPos, count)
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
        Interpolate = Strings.Replace(Interpolate, "{" & CStr(args(argIdx)) & "}", CStr(args(argIdx + 1)))
    Next argIdx
    On Error GoTo 0
    Exit Function
    
InterpolateInvalidArg:
    Err.Raise 5, "Strings.Interpolate", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function


Public Function Interpolate2(ByVal baseString As String, ByVal args As Dictionary) As String
    If baseString = vbNullString Then
        Err.Raise 5, "Strings.Interpolate", "Base string cannot be null."
    End If
    
    Dim idx As Long
    Interpolate2 = baseString
    Dim key As Variant
    On Error GoTo Interpolate2InvalidArg
    For Each key In args.Keys()
        Interpolate2 = Strings.Replace(Interpolate2, "{" & CStr(key) & "}", CStr(args.item(key)))
        idx = idx + 1
    Next key
    On Error GoTo 0
    Exit Function
    
Interpolate2InvalidArg:
    Err.Raise 5, "Strings.Interpolate", "Argument at position " & idx * 2 + 1 & " could not be converted to a string."
End Function




Public Function IsNullOrEmpty(ByVal baseString As String) As Boolean
    IsNullOrEmpty = baseString = vbNullString
End Function



Public Function IsNullOrWhiteSpace(ByVal baseString As String) As Boolean
    'equivalent to: return String.IsNullOrEmpty(value) || value.Trim().Length == 0;
    IsNullOrWhiteSpace = Strings.IsNullOrEmpty(baseString) Or Len(Strings.Trim(baseString)) = 0
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
    
    JoinBetween = Strings.Join(delimiter, subset)
End Function





Public Function LastIndexOf(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal startPos As Long = -2, Optional ByVal compare As Comparison = Comparison.Default) As Long
    If compare Mod 2 = 1 Then
        baseString = LCase$(baseString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or compare = BinaryIgnoreCase Then
        LastIndexOf = InStrRev(baseString, stringToFind, startPos + 1, vbBinaryCompare) - 1
    ElseIf compare = text Or compare = TextIngnoreCase Then
        LastIndexOf = InStrRev(baseString, stringToFind, startPos + 1, vbTextCompare) - 1
    ElseIf compare = Database Or compare = DatabaseIgnoreCase Then
        LastIndexOf = InStrRev(baseString, stringToFind, startPos + 1, vbDatabaseCompare) - 1
    Else
        LastIndexOf = InStrRev(baseString, stringToFind, startPos + 1) - 1
    End If
End Function




Public Function LastIndexOfBetween(ByVal baseString As String, ByVal stringToFind As String, ByVal startIndex As Long, ByVal Length As Long, Optional ByVal compare As Comparison = Comparison.Default) As Long
    If compare Mod 2 = 1 Then
        baseString = LCase$(baseString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or compare = BinaryIgnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(baseString, startIndex - Length, Length), stringToFind, -1, vbBinaryCompare) - 1
    ElseIf compare = text Or compare = TextIngnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(baseString, startIndex - Length, Length), stringToFind, -1, vbTextCompare) - 1
    ElseIf compare = Database Or compare = DatabaseIgnoreCase Then
        LastIndexOfBetween = InStrRev(Mid$(baseString, startIndex - Length, Length), stringToFind, -1, vbDatabaseCompare) - 1
    Else
        LastIndexOfBetween = InStrRev(Mid$(baseString, startIndex - Length, Length), stringToFind, -1) - 1
    End If
    If LastIndexOfBetween > -1 Then
        LastIndexOfBetween = LastIndexOfBetween + startIndex - Length - 1
    End If

End Function


Public Function LastIndexOfAny(ByVal baseString As String, ByRef stringsToFind() As String, Optional ByVal startPos As Long = 0, Optional ByVal Length As Long = 0) As Long
    If Length = 0 Then Length = Len(baseString)
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        LastIndexOfAny = LastIndexOfBetween(baseString, stringsToFind(idx), startPos + 1, Length)
        If LastIndexOfAny > 0 Then Exit Function
    Next idx
End Function



Public Function Left(ByVal baseString As String, ByVal count As Long) As String
    Left = VBA.Left$(baseString, count)
End Function




Public Function Length(ByVal baseString As String) As Long
    Length = Len(baseString)
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



Public Function Overwrite(ByVal baseString As String, ByVal startIndex As Long, ByVal replacementText As String) As String
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




Public Function Remove(ByVal baseString As String, ByVal startIndex As Long, Optional ByVal count As Long = 0) As String
    If startIndex < 0 Then
        Err.Raise 9, "Strings.Remove", "Start index must be greater than 0"
    ElseIf count < 0 Then
        Err.Raise 9, "Strings.Remove", "Count must be greater than 0"
    End If
    
    If count = 0 Then
        Remove = Strings.Left(baseString, count)
    Else
        Remove = Strings.Left(baseString, Len(baseString) - startIndex)
    End If
End Function



Public Function RemoveFromEndWhile(ByVal baseString As String, ByVal stringToRemove As String) As String
    If stringToRemove = vbNullString Or Len(stringToRemove) = 0 Then
        Err.Raise 9, "Strings.RemoveFromEndWhile", "String to remove must have length greater than 0 and cannot be null"
    End If
    
    Dim charCount As Long: charCount = Len(stringToRemove)
    RemoveFromEndWhile = baseString
    
    Do While Strings.Right(RemoveFromEndWhile, charCount) = stringToRemove
        RemoveFromEndWhile = Strings.Left(baseString, Len(RemoveFromEndWhile) - charCount)
    Loop
End Function



Public Function RemoveNonPrintableChars(ByVal baseString As String) As String
    'Does not remove new line characters
    Dim idx As Long
    Dim charIdx As Long
    Dim currentCharCode As Long

    RemoveNonPrintableChars = String$(Len(baseString), Chr$(0))
    For idx = 1 To Len(baseString)
        currentCharCode = AscW2(Mid$(baseString, idx, 1))
        If currentCharCode > 31 Or currentCharCode = 13 Or currentCharCode = 10 Then
            charIdx = charIdx + 1
            Mid$(RemoveNonPrintableChars, charIdx, 1) = Mid$(baseString, idx, 1)
        End If
    Next idx
    RemoveNonPrintableChars = Left$(RemoveNonPrintableChars, charIdx)
End Function



'TODO - replace & replacebetween with compare and direction
Public Function Replace(ByVal baseString As String, ByVal oldString As String, ByVal newString As String, Optional ByVal compare As Comparison = Comparison.Default) As String
    If oldString = vbNullString Then
        Err.Raise 9, "Strings.Replace", "String to replace cannot be null"
    End If
    
    Dim tempBase As String
    Dim tempOld As String
    If compare Mod 2 = 1 Then
        tempBase = LCase$(tempBase)
        tempOld = LCase$(tempOld)
    End If

    If compare = Binary Or compare = BinaryIgnoreCase Then
        Replace = VBA.Replace(baseString, oldString, newString, , , vbBinaryCompare)
    ElseIf compare = text Or compare = TextIngnoreCase Then
        Replace = VBA.Replace(baseString, oldString, newString, , , vbTextCompare)
    ElseIf compare = Database Or compare = DatabaseIgnoreCase Then
        Replace = VBA.Replace(baseString, oldString, newString, , , vbDatabaseCompare)
    Else
        Replace = VBA.Replace(baseString, oldString, newString)
    End If
    
End Function
'Public Sub Replace(ByVal oldString As String, ByVal newString As String, Optional ByVal numberOfSubstitutions As Long = -1, Optional ByVal direction As ProcessDirection = ProcessDirection.FromStart)
'    Me.ReplaceBetween oldString, newString, 0, pBuffer.endIndex_, numberOfSubstitutions, direction
'End Sub


'Public Sub ReplaceBetween(ByVal oldString As String, ByVal newString As String, ByVal startIndex As Long, ByVal count As Long, Optional ByVal numberOfSubstitutions As Long = -1, Optional ByVal direction As ProcessDirection = ProcessDirection.FromStart)
'    If oldString = vbNullString Then
'        Err.Raise 5, "StringBuilder.Replace", "String to find cannot be null"
'    ElseIf startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
'        Err.Raise 9, "StringBuilder.Replace", "Invalid startIndex"
'    ElseIf count < -1 Then
'        Err.Raise 5, "StringBuilder.Replace", "Invalid length"
'    ElseIf count = 0 Or numberOfSubstitutions = 0 Then
'        Exit Sub 'Nothing to replace
'    End If
'
'    Dim pos As Long: pos = startIndex
'    Dim countSubs As Long
'    Dim maxSubs As Long
'    If numberOfSubstitutions = -1 Then
'        maxSubs = 2147483647
'    Else
'        maxSubs = numberOfSubstitutions
'    End If
'
'    If direction = FromStart Then
'        Do While pos < count And countSubs < maxSubs
'            pos = Me.IndexOf(oldString, pos)
'            If pos = -1 Then Exit Sub 'No more replacements to make
'            Me.Remove pos, Len(oldString)
'            Me.Insert pos, newString, 1
'            countSubs = countSubs + 1
'        Loop
'    Else
'        Do While pos > pBuffer.endIndex_ - count And countSubs < maxSubs
'            pos = Me.LastIndexOf(oldString, pos)
'            If pos = -1 Then Exit Sub 'No more replacements to make
'            Me.Remove pos, Len(oldString)
'            Me.Insert pos, newString, 1
'            countSubs = countSubs + 1
'        Loop
'    End If
'End Sub





Public Function Right(ByVal baseString As String, ByVal count As Long) As String
    Right = VBA.Right$(baseString, count)
End Function



Public Function StartsWith(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    If Comparison = Comparison.Default Then
        StartsWith = Left$(baseString, Len(stringToFind)) = stringToFind
    Else
        StartsWith = IndexOfBetween(baseString, stringToFind, 0, Len(stringToFind), compare) >= 0
    End If
End Function


Public Function StartsWithAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal compare As Comparison = Comparison.Default) As Boolean
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        StartsWithAny = StartsWith(baseString, stringsToFind(idx), compare)
        If StartsWithAny Then Exit Function
    Next idx
End Function



'''===================================================================================================================================================
'''<summary>
'''Retrieves a substring from this instance. The substring starts at a specified character position and continues to the end of the string or to the specified length.
'''</summary>
'''<param name="baseString">String to manipulate.</param>
'''<param name="startIndex">The zero-based starting character position of a substring in this instance.</param>
'''<param name="count">The number of characters in the substring.</param>
'''===================================================================================================================================================
Public Function Substring(ByVal baseString As String, ByVal startIndex As Long, Optional ByVal count As Long = -1) As String
    If startIndex < 0 Or startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.Substring", "Invalid startIndex"
    ElseIf count < -1 Then
        Err.Raise 5, "Strings.Substring", "Invalid length"
    ElseIf count = 0 Then
        Substring = baseString
        Exit Function
    End If
    
    If count = -1 Then count = Len(baseString) - startIndex
    If count = -1 Then
        Substring = Mid$(baseString, startIndex + 1)
    Else
        Substring = Mid$(baseString, startIndex + 1, count)
    End If
End Function



Public Function SubstringBetween(ByVal baseString As String, ByVal firstString As String, ByVal secondString As String, Optional ByVal startIndex As Long = 0) As String
    Dim startPos As Long: startPos = Strings.IndexOf(baseString, firstString, startIndex) + Strings.Length(firstString)
    Dim endPos As Long: endPos = Strings.IndexOf(baseString, secondString, startPos)
    SubstringBetween = Strings.Substring(baseString, startPos, endPos - startPos)
End Function


Public Function Reverse(ByVal baseString As String) As String
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
    ReDim delims(0 To UBound(delimiters))
    Dim idx As Long
    For idx = 0 To UBound(delimiters)
        delims(idx) = CStr(delimiters(idx))
    Next idx

    Dim result() As String
    ReDim result(0 To Len(baseString) * 2) 'worse case scenario?

    Dim pos As Long
    Dim startPos As Long
    For idx = 0 To UBound(result)
        startPos = pos
        pos = Strings.IndexOfAny(baseString, delims, pos)
        result(idx) = Strings.Substring(baseString, startPos, pos - startPos) 'TODO - .Split indexes aren't working
        pos = pos + 1
    Next idx
    Split = result
End Function



Public Function ToCharArray(ByVal baseString As String, Optional ByVal sourceIndex As Long, Optional ByVal count As Long) As String()
    Dim Chars() As String
    ReDim Chars(0 To Len(baseString)) As String
    Dim idx As Long
    Do
        Chars(sourceIndex + idx) = Mid$(baseString, sourceIndex + idx + 1, 1)
        idx = idx + 1
    Loop While idx < count
    ToCharArray = Chars
End Function


Public Sub ToStringArray(ByRef outputArray() As String, ByRef inputArray() As Variant)
    ReDim outputArray(LBound(inputArray) To UBound(inputArray))
    Dim idx As Long
    On Error GoTo ErrorCannotConvertToString:
    For idx = 0 To UBound(outputArray)
        outputArray(idx) = CStr(inputArray(idx))
    Next idx
    On Error GoTo 0
    Exit Sub
    
ErrorCannotConvertToString:
    Err.Raise 13, "Strings.ToStringArray", "Cannot convert element at "
    On Error GoTo 0
End Sub




Public Function ToLower(ByVal baseString As String) As String
    ToLower = LCase$(baseString)
End Function




Public Function ToUpper(ByVal baseString As String) As String
    ToUpper = UCase$(baseString)
End Function


Public Function Trim(ByVal baseString As String) As String
    Trim = VBA.Trim$(baseString)
End Function


Public Function TrimEnd(ByVal baseString As String) As String
    TrimEnd = VBA.RTrim$(baseString)
End Function

Public Function TrimLeft(ByVal baseString As String) As String
    TrimLeft = VBA.LTrim$(baseString)
End Function


'/**
' * Create a max length of string and return it with extension.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @example Truncate("This is a long sentence", 10)  -> "This is..."
' */
Public Function Truncate(ByRef baseString As String, ByVal maxLength As Long, Optional ByVal extension As String = "...") As String
    If Len(baseString) <= maxLength Then
        Truncate = baseString
        Exit Function
    End If

    Truncate = Left(baseString, maxLength - Len(extension)) & extension
End Function

'from https://stackoverflow.com/a/218199/10802255
Public Function URLEncode(ByVal baseString As String, Optional ByVal spaceAsPlus As Boolean = False) As String
    Dim stringLen As Long: stringLen = Len(baseString)

    If stringLen > 0 Then
        ReDim result(stringLen) As String
        Dim idx As Long
        Dim charCode As Long
        Dim Char As String
        Dim space As String

        If spaceAsPlus Then space = "+" Else space = "%20"

        For idx = 1 To stringLen
            Char = Mid$(baseString, idx, 1)
            charCode = Asc(Char)
            Select Case charCode
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                result(idx) = Char
            Case 32
                result(idx) = space
            Case 0 To 15
                result(idx) = "%0" & Hex$(charCode)
            Case Else
                result(idx) = "%" & Hex$(charCode)
            End Select
        Next idx
        URLEncode = Join(result, vbNullString)
    End If
End Function

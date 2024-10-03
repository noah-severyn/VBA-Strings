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

'''<summary>
'''Specifies the case and sort rules to be used when comparing strings.
'''</summary>
'''<see href="https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/comparison-constants"/>
    '''<summary>
    '''Indicates the default option settings for string comparisons.
    '''</summary>
    '''<summary>
    '''Indicates that the string comparison must ignore case.
    '''</summary>
Public Enum CompareOptions
    Default = 0
    IgnoreCase = 1
End Enum



'TODO - easy way to convert between array and collection to make the _Any functions more flexible?


'''===================================================================================================================================================
'''<summary>
'''Returns a Long representing the character code corresponding to the first letter in a string.
'''</summary>
'''<param name="character">Any valid string. If the string length is greater than one, then only the first character is used for input.</param>
'''<error cref="9">String cannot be null or empty.</error>
'''<returns>A Long representing the character code corresponding to the first letter in a string.</returns>
'''<remarks>The AscW function in the built-in VBA.Strings module has a problem where it returns the correct bit pattern for an unsigned 16-bit integer which is
'''incorrect in VBA because VBA uses signed 16-bit integer. Thus, after reaching 32767 AscW will start returning negative numbers.
'''</remarks>
'''===================================================================================================================================================
'@Ignore UseMeaningfulName
Public Function AscW2(ByVal character As String) As Long
    If Len(character) > 1 Then
        character = VBA.Left$(character, 1)
    ElseIf character = vbNullString Then
        Err.Raise 9, "Strings.AscW2", "String cannot be null or empty."
    End If
        
    AscW2 = AscW(character) And &HFFFF&
End Function



'''===================================================================================================================================================
'''<summary>
'''Appends a the specified strings to the end to the end of the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="args">One or more strings to append to the base string.</param>
'''<error cref="5">An argument could not be converted to a string.</error>
'''<returns>A new string after all append operations have completed.</returns>
'''<remarks>The args parameter can be a any combination of series of single values, an array, or a collection.</remarks>
'''===================================================================================================================================================
Public Function Append(ByVal baseString As String, ParamArray args() As Variant) As String
    Append = baseString
    Dim idx As Long
    Dim idxInner As Long
    On Error GoTo AppendInvalidArg
    For idx = 0 To UBound(args)
        If IsArray(args(idx)) Then
            For idxInner = 0 To UBound(args(idx))
                Append = Append & CStr(args(idx)(idxInner))
            Next idxInner
        ElseIf TypeName(args(idx)) = "Collection" Then
            For idxInner = 1 To args(idx).count
                Append = Append & CStr(args(idx).item(idxInner))
            Next idxInner
        Else
            Append = Append & CStr(args(idx))
        End If
    Next idx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.Append", "Argument at position " & idx * 2 + 1 & " could not be converted to a string."
End Function



'''===================================================================================================================================================
'''<summary>
'''Appends the specified strings followed by the default line terminator to the end of the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="args">One or more strings to append to the base string.</param>
'''<error cref="5">An argument could not be converted to a string.</error>
'''<returns>A new string after all append operations have completed.</returns>
'''<remarks>The args parameter can be a any combination of series of single values, an array, or a collection.</remarks>
'''===================================================================================================================================================
Public Function AppendLine(ByVal baseString As String, ParamArray args() As Variant) As String
    AppendLine = baseString
    Dim idx As Long
    Dim idxInner As Long
    On Error GoTo AppendInvalidArg
    For idx = 0 To UBound(args)
        If IsArray(args(idx)) Then
            For idxInner = 0 To UBound(args(idx))
                AppendLine = AppendLine & vbNewLine & CStr(args(idx)(idxInner))
            Next idxInner
        ElseIf TypeName(args(idx)) = "Collection" Then
            For idxInner = 1 To args(idx).count
                AppendLine = AppendLine & vbNewLine & CStr(args(idx).item(idxInner))
            Next idxInner
        Else
            AppendLine = AppendLine & vbNewLine & CStr(args(idx))
        End If
    Next idx
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "Strings.AppendLine", "Argument at position " & idx * 2 + 1 & " could not be converted to a string."
End Function



'''===================================================================================================================================================
'''<summary>
'''Gets the a character at a specified position in the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="index">A position in the current string.</param>
'''<error cref="9">Index must be greater than zero.</error>
'''<returns>The character as position index.</returns>
'''===================================================================================================================================================
Public Function Chars(ByVal baseString As String, ByVal index As Long) As String
    If index < 0 Then
        Err.Raise 9, "Strings.Char", "Index must be greater than zero."
    ElseIf (index) > Len(baseString) Then
        index = Len(baseString) - 1
    End If

    Chars = Mid$(baseString, index + 1, 1)
End Function



'''===================================================================================================================================================
'''<summary>
'''Removes nonprintable characters from a string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="nonPrintable">Specify whether to remove non-printable characters.</param>
'''<param name="newLines">Specify whether to remove new line characters.</param>
'''<param name="newLineReplacement">Replacement character to insert in place of removed new line characters. Has no effect if new line characters are not removed.</param>
'''<param name="nonBreaking">Specify whether to remove non-breaking spaces.</param>
'''<param name="trimString">Specify whether to trim the string after any remove operations.</param>
'''<returns>The cleaned string.</returns>
'''===================================================================================================================================================
Public Function Clean(ByVal baseString As String, Optional ByVal nonPrintable As Boolean = True, Optional ByVal newLines As Boolean = True, Optional ByVal newLineReplacement As String = " ", Optional ByVal nonBreaking As Boolean = True, Optional ByVal trimString As Boolean = True) As String
    Clean = baseString
    If nonPrintable Then Clean = Strings.RemoveNonPrintableChars(Clean)
    If newLines Then Clean = Strings.ReplaceNewLineChars(Clean, newLineReplacement)
    If nonBreaking Then Clean = Strings.ReplaceNonBreakingSpaces(Clean)
    If trimString Then Clean = Strings.Trim(Clean)
End Function



'''===================================================================================================================================================
'''<summary>
'''Replaces the current string with vbnullstring.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>A new string containing only vbnullstring.</returns>
'''===================================================================================================================================================
Public Function Clear(ByVal baseString As String) As String
    Clear = vbNullString
End Function



'''===================================================================================================================================================
'''<summary>
'''Evaluates the arguments in order and returns the first argument that is not vbnullstring or null.
'''</summary>
'''<param name="args">One or more values to check.</param>
'''<returns>The the first non null string.</returns>
'''===================================================================================================================================================
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



'''===================================================================================================================================================
'''<summary>
'''Concatenates one or more strings together, each separated by a delimiter.
'''</summary>
'''<param name="delimiter">Any valid string.</param>
'''<param name="args">One or more strings to concatenate together.</param>
'''<error cref="5">An argument could not be converted to a string.</error>
'''<returns>The concatenated elements of args.</returns>
'''<remarks>The args parameter can be a any combination of series of single values, an array, or a collection.</remarks>
'''===================================================================================================================================================
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



'''===================================================================================================================================================
'''<summary>
'''Returns a value indicating whether a specified substring occurs within a base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to search for.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison..</param>
'''<returns>true if the stringToFind ocurrs within the baseString; otherwise, false.</returns>
'''===================================================================================================================================================
Public Function Contains(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    Contains = IndexOf(baseString, stringToFind, , compare) >= 0
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a value indicating whether a specified substring occurs within a base string after a specified position.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to search for.</param>
'''<param name="startIndex">Position to start searching from.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison..</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Start index must be less than the base string length.</error>
'''<returns>true if the stringToFind ocurrs within the baseString after the startIndex; otherwise, false.</returns>
'''===================================================================================================================================================
Public Function ContainsAfter(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal startIndex As Long, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    If startIndex < 0 Then
        Err.Raise 9, "Strings.ContainsAfter", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.ContainsAfter", "Start index must be less than the base string length."
    End If
    
    ContainsAfter = Strings.IndexOf(baseString, stringToFind, startIndex, compare) >= 0
End Function




'''===================================================================================================================================================
'''<summary>
'''Returns a value indicating whether a specified substring occurs within a base string before a specified position.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to search for.</param>
'''<param name="startIndex">Position to end searching at.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison..</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Start index must be less than the base string length.</error>
'''<returns>true if the stringToFind ocurrs within the baseString before the endIndex; otherwise, false.</returns>
'''===================================================================================================================================================
Public Function ContainsBefore(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal endIndex As Long, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    If endIndex < 0 Then
        Err.Raise 9, "Strings.ContainsBefore", "End index must be greater than zero."
    End If
    
    Dim newString As String: newString = Left(baseString, endIndex)
    ContainsBefore = Strings.IndexOf(newString, stringToFind, , compare) >= 0
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a value indicating whether any of the specified substrings occur within a base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to search for.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">Strings to find must be an array of strings.</error>
'''<returns>true if any of the strings in stringsToFind ocurr within the baseString; otherwise, false.</returns>
'''===================================================================================================================================================
Public Function ContainsAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    If Not IsArray(stringsToFind) Then
        Err.Raise 9, "Strings.ContainsAny", "Strings to find must be an array of strings."
    End If
    
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        ContainsAny = Strings.IndexOf(baseString, stringsToFind(idx), , compare) >= 0
        If ContainsAny Then Exit Function
    Next idx
End Function



'''===================================================================================================================================================
'''<summary>
'''Copies each character in the base string to a new array.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>An array of string characters from the base string.</returns>
'''===================================================================================================================================================
Public Function ConvertToCharArray(ByVal baseString As String) As Variant
    Dim idx As Long
    Dim characters As Variant
    ReDim characters(0 To Len(baseString) - 1)
    
    For idx = 1 To Len(baseString)
        characters(idx - 1) = Mid$(baseString, idx, 1)
    Next idx
    CopyToCharArray = characters
End Function



'''===================================================================================================================================================
'''<summary>
'''Copies a specified number of characters from a specified position in this instance to a specified position in an array of Unicode characters.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="sourceIndex">The index of the first character in the base string to copy.</param>
'''<param name="charArray">An array of characters to which characters in the base string are copied.</param>
'''<param name="destinationIndex" optional="true">The index in destination at which the copy operation begins.</param>
'''<param name="count" optional="true">The number of characters in this instance to copy to destination.</param>
'''===================================================================================================================================================
Public Sub CopyToCharArray(ByVal baseString As String, ByVal sourceIndex As Long, ByRef charArray As Variant, Optional ByVal destinationIndex As Long = 0, Optional ByVal count As Long = 0)
    Dim pos As Long
    If Len(baseString) - 1 + destinationIndex - 1 > UBound(charArray) - LBound(charArray) Then
        ReDim Preserve charArray(0 To Len(baseString) + destinationIndex - 1)
    End If
    
    If count <= 0 Then count = Len(baseString)
    Do
        charArray(destinationIndex + pos) = Mid$(baseString, sourceIndex + pos + 1, 1)
        pos = pos + 1
    Loop While pos < count
End Sub



'''===================================================================================================================================================
'''<summary>
'''Counts the number of ocurrences of a substring within a base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to search for.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<returns>The count of ocurrences within the base string.</returns>
'''===================================================================================================================================================
Public Function CountSubstring(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Long
    Dim locn As Long: locn = IndexOf(baseString, stringToFind, locn)

    Do While locn >= 0
        locn = IndexOf(baseString, stringToFind, locn + 1, compare)
        CountSubstring = CountSubstring + 1
    Loop
End Function



'''===================================================================================================================================================
'''<summary>
'''Creates a new string of the given length of the specified character.
'''</summary>
'''<param name="count">The length of the string to create..</param>
'''<param name="defaultChar">The character to repeat count times.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">The default character can only be one character in length.</error>
'''<returns>The created string of count length.</returns>
'''===================================================================================================================================================
Public Function Create(ByVal count As Long, Optional ByVal defaultChar As String = " ") As String
    If Not Len(defaultChar) > 1 Then
        Err.Raise 9, "Strings.Create", "The default character can only be one character in length."
    End If
    Create = Replace(space(count), " ", defaultChar)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns an empty string.
'''</summary>
'''<returns>An empty string.</returns>
'''<returns>The empty string is equivalent to vbnullstring or "".</returns>
'''===================================================================================================================================================
Public Function EmptyString() As String
    EmptyString = vbNullString
End Function



'''===================================================================================================================================================
'''<summary>
'''Determines whether the end of a base string matches the specified string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">TThe string to compare to the substring at the end of this instance.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<returns>true if stringToFind matches the end of the base string; otherwise, false.</returns>
'''===================================================================================================================================================
Public Function EndsWith(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    If CompareOptions = CompareOptions.Default Then
        EndsWith = Right$(baseString, Len(stringToFind)) = stringToFind
    Else
        EndsWith = IndexOfBetween(baseString, stringToFind, Len(baseString) - Len(stringToFind), Len(stringToFind), compare) >= 0
    End If
End Function



'''===================================================================================================================================================
'''<summary>
'''Determines whether the end of a base string matches any of the specified strings.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringsToFind">Strings to compare to the substring at the end of this instance.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<returns>true if any of stringsToFind matches the end of the base string; otherwise, false.</returns>
'''===================================================================================================================================================
Public Function EndsWithAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        EndsWithAny = EndsWith(baseString, stringsToFind(idx), compare)
        If EndsWithAny Then Exit Function
    Next idx
End Function



'''===================================================================================================================================================
'''<summary>
'''Determine whether two strings have the same value.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="compareString">Any valid string to compare.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<returns>true if the strings are identical; otherwise, false. If compareString is null, the method returns false.</returns>
'''===================================================================================================================================================
Public Function Equals(ByVal baseString As String, ByVal compareString As String, ByVal compare As CompareOptions) As Boolean
    If compareString = vbNullString Then
        Equals = False
        Exit Function
    End If
    
    If compare Mod 2 = 1 Then
        baseString = LCase$(baseString)
        compareString = LCase$(compareString)
    End If
    
    Equals = baseString = compareString
End Function



'''===================================================================================================================================================
'''<summary>
'''Determine whether a string has the same value as any other string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<param name="compareStrings">One or more stirngs to compare to.</param>
'''<returns>true if any of the strings in compareStrings are identical to the base string.; otherwise, false.</returns>
'''<remarks>compareStrings can be a series of strings, a collection, an array, or any combination of the three.</remarks>
'''===================================================================================================================================================
Public Function EqualsAny(ByVal baseString As String, ByVal compare As CompareOptions, ParamArray compareStrings() As Variant) As Boolean
    Dim idx As Long
    Dim idxInner As Long
    For idx = LBound(compareStrings) To UBound(compareStrings)
        If TypeName(comarestrings(idx)) = "Collection" Then
            For idxInner = 1 To compareStrings(idx).count
                EqualsAny = Strings.Equals(baseString, compare, compareStrings(idx).item(idxInner))
                If EqualsAny = True Then Exit Function
            Next idxInner
        ElseIf IsArray(compareStrings(idx)) Then
            For idxInner = LBound(compareStrings(idx)) To UBound(compareStrings(idx))
                EqualsAny = Strings.Equals(baseString, compare, compareStrings(idx)(idxInner))
                If EqualsAny = True Then Exit Function
            Next idxInner
        Else
            EqualsAny = Strings.Equals(baseString, compare, compareStrings(idx))
            If EqualsAny = True Then Exit Function
        End If
    Next idx
End Function



'''===================================================================================================================================================
'''<summary>
'''Reports the zero-based index of the first occurrence of the specified string in the base string object.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to seek.</param>
'''<param name="startIndex">The search starting position.</param>
'''<param name="count">The number of character positions to examine.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Start index must be less than the base string length.</error>
'''<returns>The zero-based index position of stringToFind from the start of the base string if that string is found, or -1 if it is not. If stringToFind is null, the return value is startIndex.</returns>
'''===================================================================================================================================================
Public Function IndexOf(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal startIndex As Long = 0, Optional ByVal count As Long = -1, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Long
    Dim startPos As Long
    If startIndex < 0 Then
        Err.Raise 9, "Strings.IndexOf", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.IndexOf", "Start index must be less than the base string length."
    ElseIf stringToFind = vbNullString Then
        IndexOf = startIndex
        Exit Function
    ElseIf count = -1 Then
        count = Len(baseString) - startIndex
    ElseIf startIndex + count > Len(baseString) + 1 Then
        count = Len(baseString) - startIndex
    End If
    
    Dim vbComp As VbCompareMethod
    If compare = CompareOptions.IgnoreCase Then
        vbComp = vbTextCompare
    Else
        vbComp = vbBinaryCompare
    End If
    startPos = startIndex + 1 'Adjustment for the 1-based strings in VBA
    
    IndexOf = InStr(1, Mid$(baseString, startPos, count), stringToFind, vbComp) - 1
    If IndexOf <> -1 Then IndexOf = IndexOf + startPos
End Function



'''===================================================================================================================================================
'''<summary>
'''Reports the zero-based index of the first occurrence in the base string of any string in a specified array of strings. The search starts at a specified character position and examines a specified number of character positions.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringsToFind">An array or collection of strings to seek.</param>
'''<param name="startIndex">The search starting position.</param>
'''<param name="count">The number of character positions to examine.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Start index must be less than the base string length.</error>
'''<returns>The zero-based index position of the first occurrence in this instance where any string in stringsToFind was found; -1 if no string in stringsToFind was found.</returns>
'''===================================================================================================================================================
Public Function IndexOfAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal startIndex As Long = 0, Optional ByVal Count As Long = 0, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Long
    Dim maxLoops As Long
    If startIndex < 0 Then
        Err.Raise 9, "Strings.IndexOfAny", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.IndexOfAny", "Start index must be less than the base string length."
    ElseIf Not VBA.IsArray(stringsToFind) And Not TypeName(stringsToFind) = "Collection" Then
        Err.Raise 9, "Strings.IndexOfAny", "The 'stringsToFind' parameter is not an array or collection."
    ElseIf VBA.IsArray(stringsToFind) Then
        If UBound(stringsToFind) - LBound(stringsToFind) + 1 = 0 Then
            IndexOfAny = startIndex
            Exit Function
        End If
        maxLoops = UBound(stringsToFind) - LBound(stringsToFind) + 1
    ElseIf TypeName(stringsToFind) = "Collection" Then
        If stringsToFind.Count = 0 Then
            IndexOfAny = startIndex
            Exit Function
        End If
        maxLoops = stringsToFind.Count
    End If
    
    If Count = -1 Then
        Count = Len(baseString) - startIndex
    ElseIf Count = 0 Then
        Count = Len(baseString)
    End If
    If startIndex + Count > Len(baseString) Then
        Count = Len(baseString) - startIndex + 1
    End If
    
    Dim idx As Long
    Dim result As Long
    IndexOfAny = 99999999
    For idx = 0 To maxLoops - 1
        result = Strings.IndexOf(baseString, stringsToFind(idx - CLng(TypeName(stringsToFind) = "Collection")), startIndex, Count, compare)
        If result < IndexOfAny And result <> -1 Then
            IndexOfAny = result
        End If
    Next idx
    If IndexOfAny = 99999999 Then IndexOfAny = -1
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which a specified string is inserted at a specified index position in the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="startIndex">The zero-based index position of the insertion.</param>
'''<param name="stringToInsert">The string to insert.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Start index must be less than the base string length.</error>
'''<returns>A new string that is equivalent to this instance, but with value inserted at position startIndex.</returns>
'''===================================================================================================================================================
Public Function Insert(ByVal baseString As String, ByVal startIndex As Long, ByVal stringToInsert As String) As String
    If startIndex < 0 Then
        Err.Raise 9, "Strings.IndexOf", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.IndexOf", "Start index must be less than the base string length."
    End If
    
    Insert = Left$(baseString, startIndex) & stringToInsert & Right$(baseString, Len(baseString) - startIndex)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which the special marked keys and values are replaced their associated values.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="args">A series of key and value strings to use to substitute.</param>
'''<error cref="5">Base string cannot be null.</error>
'''<error cref="5">Invalid number of parameters. The interpolation parameters must be provided in pairs.</error>
'''<returns>A new string with substitued values.</returns>
'''===================================================================================================================================================
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



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which the special marked keys and values are replaced their associated values.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="args">A dictionary of key and value strings to use to substitute.</param>
'''<error cref="5">Base string cannot be null.</error>
'''<returns>A new string with substitued values.</returns>
'''===================================================================================================================================================
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


'''===========================================================================================================================================================================================
'''<summary>Indicates whether the specified character is alphanumeric.</summary>
'''<param name="character" type="String">Character to evaluate. If this string has length greater than 1 then only the first character is examined</param>
'''<returns type="Boolean">TRUE if character is alphanumeric; FALSE otherwise</returns>
'''<remarks>
'''Currently only valid for ASCII characters 0-127
'''</remarks>
'''===========================================================================================================================================================================================
Public Function IsAlphanumeric(ByVal character As String) As Boolean
    If character = vbNullString Then
        IsAlphanumeric = False
        Exit Function
    End If
    Dim isnumeric As Boolean: isnumeric = Asc(character) >= 48 And Asc(character) <= 57
    Dim isalphabetic As Boolean: isalphabetic = (Asc(character) >= 65 And Asc(character) <= 90) Or (Asc(character) >= 97 And Asc(character) <= 122)

    IsAlphanumeric = isnumeric Or isalphabetic
End Function



'''===================================================================================================================================================
'''<summary>
'''Indicates whether the specified string is null (vbNullString).
'''</summary>
'''<param name="baseString">The string to test.</param>
'''<returns>TRUE if the value parameter is null or an empty string; otherwise, FALSE.</returns>
'''===================================================================================================================================================
Public Function IsNullOrEmpty(ByVal baseString As String) As Boolean
    IsNullOrEmpty = baseString = vbNullString
End Function



'''===================================================================================================================================================
'''<summary>
'''Indicates whether the specified string is null (vbNullString) or consists only of white-space characters.
'''</summary>
'''<param name="baseString">The string to test.</param>
'''<returns>TRUE if the value parameter is null or an empty string or white space; otherwise, FALSE.</returns>
'''<remarks>This is a convenience function equivalent to <c>Strings.IsNullOrEmpty(baseString) Or Strings.Trim(basestring).Length = 0</c>.</remarks>
'''===================================================================================================================================================
Public Function IsNullOrWhiteSpace(ByVal baseString As String) As Boolean
    IsNullOrWhiteSpace = Strings.IsNullOrEmpty(baseString) Or Len(Strings.Trim(baseString)) = 0
End Function



'''===================================================================================================================================================
'''<summary>
'''Concatenates all the elements of a string array, using the specified separator between each element.
'''</summary>
'''<param name="separator">The string to use as a separator. separator is included in the returned string only if value has more than one element.</param>
'''<param name="stringsToJoin">An array that contains the elements to concatenate.</param>
'''<returns>A string that consists of the elements in value delimited by the separator string.</returns>
'''===================================================================================================================================================
Public Function Join(ByVal separator As String, ByVal stringsToJoin As Variant) As String
    If VBA.IsArray(stringsToJoin) Then
        Join = VBA.Join(stringsToJoin, separator)
    Else
        Dim idx As Long
        For idx = 1 To stringsToJoin.count
            Join = Join & stringsToJoin(idx) & separator
        Next idx
        Join = Left$(Join, Len(Join) - 1)
    End If
End Function



'''===================================================================================================================================================
'''<summary>
'''Concatenates an array of strings, using the specified separator between each member, starting with the element in value located at the startIndex position, and concatenating up to count elements.
'''</summary>
'''<param name="separator">The string to use as a separator. separator is included in the returned string only if value has more than one element.</param>
'''<param name="stringsToJoin">An array that contains the elements to concatenate.</param>
'''<param name="startIndex">The first item in stringsToJoin to concatenate.</param>
'''<param name="count">The number of elements from stringsToJoin to concatenate, starting with the element in the startIndex position.</param>
'''<returns>A string that consists of count elements of stringsToJoin starting at startIndex delimited by the separator character.</returns>
'''===================================================================================================================================================
Public Function JoinBetween(ByVal separator As String, ByVal stringsToJoin As Variant, ByVal startIndex As Long, ByVal count As Long) As String
    If startIndex < 0 Then
        Err.Raise 9, "Strings.JoinBetween", "The startIndex must be greater than zero."
    ElseIf count < 0 Then
        Err.Raise 5, "Strings.JoinBetween", "The count must be greater than zero"
    ElseIf count = 0 Then
        Exit Function
    End If
    
    Dim subset As Variant
    ReDim subset(0 To count - 1)
    Dim idx As Long
    For idx = 0 To count - 1
        subset(idx) = stringsToJoin(idx + startIndex)
    Next idx
    
    JoinBetween = Strings.Join(separator, subset)
End Function



'''===================================================================================================================================================
'''<summary>
'''Reports the zero-based index position of the last occurrence of a specified string within a string. The search starts at a specified character position
'''and proceeds backward toward the beginning of the string for the specified number of character positions. A parameter specifies the type of comparison
'''to perform when searching for the specified string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to seek.</param>
'''<param name="startIndex">The search starting position.</param>
'''<param name="length">The number of character positions to examine.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Start index must be less than the base string length.</error>
'''<returns>The zero-based starting index position of the stringToFind if that string is found, or -1 if it is not found or if the base string is null</returns>
'''===================================================================================================================================================
Public Function LastIndexOf(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal startIndex As Long = -2, Optional ByVal count As Long = -1, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Long
    If startIndex < 0 And startIndex <> -2 Then
        Err.Raise 9, "Strings.IndexOf", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.IndexOf", "Start index must be less than the base string length."
    ElseIf stringToFind = vbNullString Then
        LastIndexOf = -1
        Exit Function
    ElseIf count = -1 Or startIndex + count > Len(baseString) Then
        count = Len(baseString) - startIndex
    End If
    
    Dim vbComp As VbCompareMethod
    If compare = CompareOptions.IgnoreCase Then
        vbComp = vbTextCompare
    Else
        vbComp = vbBinaryCompare
    End If
    
    Dim substr As String
    If startIndex <> -2 Then
        substr = Mid$(baseString, startIndex - count, count)
    Else
        substr = baseString
    End If
    
    LastIndexOf = InStrRev(substr, stringToFind, startIndex + 1, vbComp) - 1
End Function



'''===================================================================================================================================================
'''<summary>
'''Reports the zero-based index position of the last occurrence of any of the specified strings within a string. The search starts at a specified character
'''position and proceeds backward toward the beginning of the string for the specified number of character positions. A parameter specifies the type of
'''comparison to perform when searching for the specified string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">An array of strings to seek.</param>
'''<param name="startIndex">The search starting position.</param>
'''<param name="count">The number of character positions to examine.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Start index must be less than the base string length.</error>
'''<returns>The zero-based starting index position of any element in stringsToFind if that string is found, or -1 if it is not found or if the base string is null</returns>
'''===================================================================================================================================================
Public Function LastIndexOfAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal startIndex As Long = 0, Optional ByVal count As Long = 0, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Long
    Dim maxLoops As Long
    If startIndex < 0 Then
        Err.Raise 9, "Strings.LastIndexOfAny", "Start index must be greater than zero."
    ElseIf startIndex > Len(baseString) Then
        Err.Raise 9, "Strings.LastIndexOfAny", "Start index must be less than the base string length."
    ElseIf Not VBA.IsArray(stringsToFind) And Not TypeName(stringsToFind) = "Collection" Then
        Err.Raise 9, "Strings.LastIndexOfAny", "The 'stringsToFind' parameter is not an array or collection."
    ElseIf VBA.IsArray(stringsToFind) Then
        If UBound(stringsToFind) - LBound(stringsToFind) + 1 = 0 Then
            LastIndexOfAny = startIndex
            Exit Function
        End If
        maxLoops = UBound(stringsToFind) - LBound(stringsToFind) + 1
    ElseIf TypeName(stringsToFind) = "Collection" Then
        If stringsToFind.count = 0 Then
            LastIndexOfAny = startIndex
            Exit Function
        End If
        maxLoops = stringsToFind.count
    End If
    
    If count = -1 Then
        count = Len(baseString) - startIndex
    ElseIf count = 0 Then
        count = Len(baseString)
    End If
    If startIndex + count > Len(baseString) Then
        count = Len(baseString) - startIndex + 1
    End If
    
    Dim idx As Long
    Dim result As Long
    LastIndexOfAny = -1
    For idx = 0 To maxLoops - 1
        result = LastIndexOf(baseString, stringsToFind(idx), startIndex + 1, count, compare)
        If result > LastIndexOfAny And result <> -1 Then
            LastIndexOfAny = result
        End If
    Next idx
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a string containing a specified number of characters from the left side of the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="count">The number of characters to return.</param>
'''<returns>A new string with count characters from the start of the base string.</returns>
'''===================================================================================================================================================
Public Function Left(ByVal baseString As String, ByVal count As Long) As String
    Left = VBA.Left$(baseString, count)
End Function



'''===================================================================================================================================================
'''<summary>
'''Gets the number of characters in the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>The length of the string.</returns>
'''===================================================================================================================================================
Public Function Length(ByVal baseString As String) As Long
    Length = Len(baseString)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns the Levenshtein distance between two strings. Informally, the Levenshtein distance between two words is the minimum number of single-character edits (insertions, deletions or substitutions) required to change one word into the other
'''</summary>
'''<param name="firstString">Any valid string.</param>
'''<param name="secondString">Any valid string.</param>
'''<returns>The Levenshtein distance between the two strings.</returns>
'''<remarks>Implementation courtesy of Robert Todar, <see href="https://github.com/todar/VBA-Strings/blob/master/StringFunctions.bas#L50">VBA-Strings</see>.</remarks>
'''===================================================================================================================================================
Public Function LevenshteinDistance(ByVal firstString As String, ByVal secondString As String) As Double
    Dim firstLength As Long: firstLength = Len(firstString)
    Dim secondLength As Long: secondLength = Len(secondString)

    'Prepare distance array matrix with the proper indexes
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

    'Outer loop is for the first string
    For index = 1 To firstLength
        'Inner loop is for the second string
        For innerIndex = 1 To secondLength
            'Character matches exactly
            If Mid$(firstString, index, 1) = Mid$(secondString, innerIndex, 1) Then
                distance(index, innerIndex) = distance(index - 1, innerIndex - 1)

            'Character is off, offset the matrix by the appropriate number.
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

    'Levenshtein is the last index of the array.
    LevenshteinDistance = distance(firstLength, secondLength)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns the similarity between to strings as a percentage. The similarity is measured from the Levenshtein distance.
'''</summary>
'''<param name="firstString">Any valid string.</param>
'''<param name="secondString">Any valid string.</param>
'''<returns>The Levenshtein distance between the two strings.</returns>
'''<remarks>Implementation courtesy of Robert Todar, <see href="https://github.com/todar/VBA-Strings/blob/master/StringFunctions.bas#L34">VBA-Strings</see>.</remarks>
'''===================================================================================================================================================
Public Function MeasureSimilarity(ByVal firstString As String, ByVal secondString As String) As Double
    Dim levenshtein As Double
    levenshtein = LevenshteinDistance(firstString, secondString)

    'Convert levenshtein into a percentage(0 to 100)
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



'''===================================================================================================================================================
'''<summary>
'''Returns a new string that right-aligns the characters in this instance by padding them on the left with a specified character, for a specified total length.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="totalWidth">The number of characters in the resulting string, equal to the number of original characters plus any additional padding characters.</param>
'''<param name="paddingChar">A character to use as padding.</param>
'''<error cref="9">Padding character can only be one character in length.</error>
'''<error cref="9">Padding character cannot be null.</error>
'''<error cref="9">Total width cannot be less than zero.</error>
'''<returns>
'''A new string that is equivalent to the base string, but right-aligned and padded on the left with as many paddingChar characters as needed to create
'''a length of totalWidth. However, if totalWidth is less than or equal to the length of the base string, the method returns a the base string.
'''</returns>
'''===================================================================================================================================================
Public Function PadLeft(ByVal baseString As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String = " ") As String
    If Len(paddingChar) > 1 Then
        Err.Raise 9, "Strings.PadLeft", "Padding character can only be one character in length."
    ElseIf paddingChar = vbNullString Then
        Err.Raise 9, "Strings.PadLeft", "Padding character cannot be null."
    ElseIf totalWidth < 0 Then
        Err.Raise 9, "Strings.PadLeft", "Total width cannot be less than zero."
    ElseIf totalWidth < Len(baseString) Then
        PadLeft = baseString
        Exit Function
    End If
    
    PadLeft = VBA.String$(totalWidth - Len(baseString), paddingChar) & baseString
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string that left-aligns the characters in this instance by padding them on the right with a specified character, for a specified total length.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="totalWidth">The number of characters in the resulting string, equal to the number of original characters plus any additional padding characters.</param>
'''<param name="paddingChar">A character to use as padding.</param>
'''<error cref="9">Padding character can only be one character in length.</error>
'''<error cref="9">Padding character cannot be null.</error>
'''<error cref="9">Total width cannot be less than zero.</error>
'''<returns>
'''A new string that is equivalent to the base string, but left-aligned and padded on the right with as many paddingChar characters as needed to create
'''a length of totalWidth. However, if totalWidth is less than or equal to the length of the base string, the method returns a the base string.
'''</returns>
'''===================================================================================================================================================
Public Function PadRight(ByVal baseString As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String = " ") As String
    If Len(paddingChar) > 1 Then
        Err.Raise 9, "Strings.PadRight", "Padding character can only be one character in length"
    ElseIf paddingChar = vbNullString Then
        Err.Raise 9, "Strings.PadRight", "Padding character cannot be null"
    ElseIf totalWidth < 0 Then
        Err.Raise 9, "Strings.PadLeft", "Total width cannot be less than zero."
    ElseIf totalWidth < Len(baseString) Then
        PadRight = baseString
        Exit Function
    End If
    
    PadRight = baseString & VBA.String$(totalWidth - Len(baseString), paddingChar)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which a specified number of characters in the base string beginning at a specified position have been deleted.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="startIndex">The zero-based position to begin deleting characters.</param>
'''<param name="count">The number of characters to delete.</param>
'''<error cref="9">Start index must be greater than 0.</error>
'''<error cref="9">Count must be greater than 0.</error>
'''<returns>A new string that is equivalent to the base string except for the removed characters.</returns>
'''<remarks>If the combination of startIndex + count is greater than the length of base string, the count is adjusted to the length of the base string minus startIndex</remarks>
'''===================================================================================================================================================
Public Function Remove(ByVal baseString As String, ByVal startIndex As Long, Optional ByVal count As Long = 0) As String
    If startIndex < 0 Then
        Err.Raise 9, "Strings.Remove", "Start index must be greater than 0."
    ElseIf count < 0 Then
        Err.Raise 9, "Strings.Remove", "Count must be greater than 0."
    ElseIf startIndex + count > Len(baseString) Then
        count = Len(baseString) - startIndex
    End If
    
    If count = 0 Then
        Remove = Strings.Left(baseString, count)
    Else
        Remove = Strings.Left(baseString, Len(baseString) - startIndex)
    End If
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which all consecutive instances of the specified character are removed from the end of the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToRemove">The string to remove.</param>
'''<error cref="9">String to remove must have length greater than 0 and cannot be null.</error>
'''<returns>A new string that is equivalent to the base string except for the removed characters.</returns>
'''===================================================================================================================================================
Public Function RemoveFromEndWhile(ByVal baseString As String, ByVal stringToRemove As String) As String
    If stringToRemove = vbNullString Or Len(stringToRemove) = 0 Then
        Err.Raise 9, "Strings.RemoveFromEndWhile", "String to remove must have length greater than 0 and cannot be null."
    End If
    
    Dim charCount As Long: charCount = Len(stringToRemove)
    RemoveFromEndWhile = baseString
    
    Do While Strings.Right(RemoveFromEndWhile, charCount) = stringToRemove
        RemoveFromEndWhile = Strings.Left(baseString, Len(RemoveFromEndWhile) - charCount)
    Loop
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which all non printable characters are removed from of the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>A new string that is equivalent to the base string except for the removed characters.</returns>
'''<remarks>New line characters are not removed</remarks>
'''===================================================================================================================================================
Public Function RemoveNonPrintableChars(ByVal baseString As String) As String
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



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which all occurrences of a specified string are replaced with another specified string, using the provided comparison type.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="oldString">The string to be replaced.</param>
'''<param name="newString">The string to replace all occurrences of oldValue.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">String to replace cannot be null.</error>
'''<returns>A string that is equivalent to the current string except that all instances of oldString are replaced with newString. If oldString is not found in the current instance, the method returns the current instance unchanged.</returns>
'''===================================================================================================================================================
Public Function Replace(ByVal baseString As String, ByVal oldString As String, ByVal newString As String, Optional ByVal compare As CompareOptions = CompareOptions.Default) As String
    If oldString = vbNullString Then
        Err.Raise 9, "Strings.Replace", "String to replace cannot be null."
    End If
    
    Dim vbComp As VbCompareMethod
    If compare = CompareOptions.IgnoreCase Then
        vbComp = vbTextCompare
    Else
        vbComp = vbBinaryCompare
    End If
    Replace = VBA.Replace(baseString, oldString, newString, , , vbComp)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which all occurrences of all the specified strings are replaced with another specified string, using the provided comparison type.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="replacementPairs">Collection of values to replace. Key = oldString (the string to be replaced), Value = newString (the string to replace all occurrences of oldValue)</param>
'''<param name="newString"></param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<error cref="9">String to replace cannot be null.</error>
'''<returns>A string that is equivalent to the current string except that all instances of oldString are replaced with newString. If oldString is not found in the current instance, the method returns the current instance unchanged.</returns>
'''===================================================================================================================================================
Public Function ReplaceAll(ByVal baseString As String, ByVal replacementPairs As Scripting.Dictionary, Optional ByVal compare As CompareOptions = CompareOptions.Default) As String
    Dim result As String: result = baseString
    
    Dim vbComp As VbCompareMethod
    If compare = CompareOptions.IgnoreCase Then
        vbComp = vbTextCompare
    Else
        vbComp = vbBinaryCompare
    End If
    
    Dim idx As Long
    Dim oldString As String
    Dim newString As String
    For idx = 0 To replacementPairs.Count - 1
        oldString = replacementPairs.Keys(idx)
        newString = replacementPairs.Items(idx)
        If oldString = vbNullString Then
            Err.Raise 9, "Strings.Replace", "String to replace cannot be null."
        End If
        
        result = VBA.Replace(result, oldString, newString, , , vbComp)
    Next idx
    ReplaceAll = result
End Function



'TODO - replace & replacebetween with compare and direction
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




'''===================================================================================================================================================
'''<summary>
'''Returns a string containing a specified number of characters from the right side of the base string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="count">The number of characters to return.</param>
'''<returns>A new string with count characters from the end of the base string.</returns>
'''===================================================================================================================================================
Public Function Right(ByVal baseString As String, ByVal count As Long) As String
    Right = VBA.Right$(baseString, count)
End Function



'''===================================================================================================================================================
'''<summary>
'''Determines whether the beginning of the base string matches the specified string when compared using the specified comparison option.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringToFind">The string to compare.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<returns>true if this instance begins with value; otherwise, false.</returns>
'''===================================================================================================================================================
Public Function StartsWith(ByVal baseString As String, ByVal stringToFind As String, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    If CompareOptions = CompareOptions.Default Then
        StartsWith = Left$(baseString, Len(stringToFind)) = stringToFind
    Else
        StartsWith = IndexOfBetween(baseString, stringToFind, 0, Len(stringToFind), compare) >= 0
    End If
End Function



'''===================================================================================================================================================
'''<summary>
'''Determines whether the beginning of the base string matches any of the specified strings when compared using the specified comparison option.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="stringsToFind">The string to compare.</param>
'''<param name="compare">One of the enumeration values that specifies the rules to use in the comparison.</param>
'''<returns>true if this instance begins with value; otherwise, false.</returns>
'''<remarks>stringsToFind can be an array or collection</remarks>
'''===================================================================================================================================================
Public Function StartsWithAny(ByVal baseString As String, ByVal stringsToFind As Variant, Optional ByVal compare As CompareOptions = CompareOptions.Default) As Boolean
    Dim idx As Long
    If IsArray(stringsToFind) Then
        For idx = 0 To UBound(stringsToFind)
            StartsWithAny = StartsWith(baseString, stringsToFind(idx), compare)
            If StartsWithAny Then Exit Function
        Next idx
    ElseIf TypeName(stringsToFind) = "Collection" Then
        For idx = 1 To stringsToFind.count
            StartsWithAny = StartsWith(baseString, stringsToFind.item(idx), compare)
            If StartsWithAny Then Exit Function
        Next idx
    End If
End Function



'''===================================================================================================================================================
'''<summary>
'''Retrieves a substring from this instance. The substring starts at a specified character position and continues to the end of the string or to the specified length.
'''</summary>
'''<param name="baseString">String to manipulate.</param>
'''<param name="startIndex">The zero-based starting character position of a substring in this instance.</param>
'''<param name="count">The number of characters in the substring.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Count must be greater than zero.</error>
'''<returns>A string that is equivalent to the substring that begins at startIndex of baseString, the baseString if count is 0, or vbnullstring if startIndex is equal to the length of this instance.</returns>
'''===================================================================================================================================================
Public Function Substring(ByVal baseString As String, ByVal startIndex As Long, Optional ByVal count As Long = -1) As String
    If startIndex < 0 Then
        Err.Raise 9, "Strings.Substring", "Start index must be greater than zero."
    ElseIf count < -1 Then
        Err.Raise 5, "Strings.Substring", "Count must be greater than zero."
    ElseIf count = 0 Then
        Substring = baseString
        Exit Function
    ElseIf startIndex = Len(baseString) Then
        Substring = vbNullString
        Exit Function
    End If
    
    If count = -1 Then
        Substring = VBA.Mid$(baseString, startIndex + 1)
    Else
        Substring = VBA.Mid$(baseString, startIndex + 1, count)
    End If
End Function



'''===================================================================================================================================================
'''<summary>
'''Retrieves a substring from this instance. The substring starts at the end of firstString and ends at the start location of secondString, optionally beginning the search for firstString at startIndex.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="firstString">String marking the start of the substring.</param>
'''<param name="secondString">String marking the end of the substring</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Value of first string was not found in the base string.</error>
'''<error cref="9">Value of second string was not found in the base string.</error>
'''<param name="startIndex">The zero-based starting character position to begin searching for firstString at.</param>
'''===================================================================================================================================================
Public Function SubstringBetween(ByVal baseString As String, ByVal firstString As String, ByVal secondString As String, Optional ByVal startIndex As Long = 0) As String
    Dim startPos As Long: startPos = Strings.IndexOf(baseString, firstString, startIndex) + Len(firstString)
    If startIndex < 0 Then
        Err.Raise 9, "Strings.Substring", "Start index must be greater than zero."
    End If
    If startPos = -1 Then
        Err.Raise 9, "Strings.SubstringBetween", "Value of first string was not found in the base string."
    End If
    Dim endPos As Long: endPos = Strings.IndexOf(baseString, secondString, startPos)
    If endPos = -1 Then
        Err.Raise 9, "Strings.SubstringBetween", "Value of second string was not found in the base string."
    End If
    
    SubstringBetween = VBA.Mid$(baseString, startPos, endPos - startPos)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a string in which the character order of a the base string is reversed.
'''</summary>
'''<param name="baseString">String to reverse.</param>
'''<returns>A reversed string.</returns>
'''===================================================================================================================================================
Public Function Reverse(ByVal baseString As String) As String
    Reverse = VBA.StrReverse$(baseString)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which all non-breaking spaces have been replaced with a standard space character.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>A new string that is equivalent to the base string except for the replaced characters.</returns>
'''===================================================================================================================================================
Public Function ReplaceNonBreakingSpaces(ByVal baseString As String) As String
    ReplaceNonBreakingSpaces = Replace(baseString, Chr$(160), " ")
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string in which all new line characters have been replaced with the specified character.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="replacement">String to substitute in place of new line characters.</param>
'''<returns>A new string that is equivalent to the base string except for the replaced characters.</returns>
'''<remarks>vbCrLf [chr(13)+chr(10)], vbCr [chr(10)], and vbLf [chr(13)] are replaced with the specified replacement or vbnNewLine, a platform-specific
'''new line character, which returns whichever new line combination is appropriate for current platform</remarks>
'''===================================================================================================================================================
Public Function ReplaceNewLineChars(ByVal baseString As String, Optional ByVal replacement As String = vbNewLine) As String
    ReplaceNewLineChars = VBA.Replace$(baseString, vbCrLf, replacement)
    ReplaceNewLineChars = VBA.Replace$(ReplaceNewLineChars, vbCr, replacement)
    ReplaceNewLineChars = VBA.Replace$(ReplaceNewLineChars, vbLf, replacement)
End Function



'''===================================================================================================================================================
'''<summary>
'''Splits a string into a maximum number of substrings based on specified delimiting characters and, optionally, options.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="includeEmptyEntries">Specify to include array elements that contain an empty string from the result.</param>
'''<param name="trimEntries">Specify to trim white-space characters from each substring in the result. If includeEmptyEntries and trimEntries are specified together, then substrings that consist only of white-space characters are also removed from the result.</param>
'''<param name="separators">An array of characters that delimit the substrings in this string, an empty array that contains no delimiters, or null.</param>
'''<returns>An array that contains the substrings in this string that are delimited by one or more characters in separator.</returns>
'''<remarks></remarks>
'TODO - .Split - add all of the remarks at https://learn.microsoft.com/en-us/dotnet/api/system.string.split?view=net-7.0#system-string-split(system-char()-system-int32-system-stringsplitoptions)
'''===================================================================================================================================================
Public Function Split(ByVal baseString As String, ByVal includeEmptyEntries As Boolean, ByVal trimEntries As Boolean, ParamArray separators() As Variant) As Variant
    'TODO .Split implement includeEmptyEntries and trimEntries options
    Dim delims As Variant
    ReDim delims(0 To UBound(separators))
    Dim idx As Long
    For idx = 0 To UBound(separators)
        delims(idx) = CStr(separators(idx))
    Next idx

    Dim result As Variant
    ReDim result(0 To Len(baseString) * 2) 'worse case scenario?

    Dim pos As Long
    Dim startPos As Long
    For idx = 0 To UBound(result)
        
        pos = Strings.IndexOfAny(baseString, delims, startPos)
        result(idx) = Strings.Substring(baseString, startPos, pos)
        startPos = startPos + pos + 1
        If pos = -1 Then Exit For
    Next idx
    ReDim Preserve result(0 To idx)
    Split = result
End Function



'''===================================================================================================================================================
'''<summary>
'''Copies the characters in a specified substring of the base string to an array.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="startIndex">The starting position of a substring in this instance.</param>
'''<param name="count">The length of the substring in this instance.</param>
'''<returns>A string array whose elements are the length number of characters in the base string starting from character position startIndex.</returns>
'''===================================================================================================================================================
Public Function ToCharArray(ByVal baseString As String, Optional ByVal startIndex As Long = 0, Optional ByVal count As Long = -1) As Variant
    If count = -1 Then
        count = Len(baseString) - startIndex
    End If
    
    Dim characters As Variant
    ReDim characters(0 To count - 1)
    Dim idx As Long
    Do
        characters(idx) = Mid$(baseString, startIndex + idx + 1, 1)
        idx = idx + 1
    Loop While idx < count
    ToCharArray = characters
End Function



'''===================================================================================================================================================
'''<summary>
'''Converts an array of variants to an array of strings.
'''</summary>
'''<param name="inputArray">Array of variants to convert.</param>
'''<error cref="13">An element of the input array cannot be converted to a string.</error>
'''<returns>A string array whose elements the string equivalents of the values of inputArray.</returns>
'''===================================================================================================================================================
Public Function ToStringArray(ByVal inputArray As Variant) As String()
    Dim outarray() As String
    ReDim outarray(LBound(inputArray) To UBound(inputArray))
    Dim idx As Long
    On Error GoTo ErrorCannotConvertToString:
    For idx = 0 To UBound(outarray)
        outarray(idx) = CStr(inputArray(idx))
    Next idx
    On Error GoTo 0
    Exit Function
    
ErrorCannotConvertToString:
    Err.Raise 13, "Strings.ToStringArray", "Cannot convert element at " & idx
    On Error GoTo 0
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string converted to lowercase.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>A string in lowercase.</returns>
'''===================================================================================================================================================
Public Function ToLower(ByVal baseString As String) As String
    ToLower = LCase$(baseString)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string converted to uppercase.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>A string in uppercase.</returns>
'''===================================================================================================================================================
Public Function ToUpper(ByVal baseString As String) As String
    ToUpper = UCase$(baseString)
End Function



'''===================================================================================================================================================
'''<summary>
'''Removes all leading and trailing white-space characters from the specified string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>The string that remains after all white-space characters are removed from the start and end of the specified string. If no characters can be trimmed, the base string is unchanged.</returns>
'''===================================================================================================================================================
Public Function Trim(ByVal baseString As String) As String
    Trim = VBA.Trim$(baseString)
End Function



'''===================================================================================================================================================
'''<summary>
'''Removes all trailing white-space characters from the specified string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>The string that remains after all white-space characters are removed from the end of the specified string. If no characters can be trimmed, the base string is unchanged.</returns>
'''===================================================================================================================================================
Public Function TrimEnd(ByVal baseString As String) As String
    TrimEnd = VBA.RTrim$(baseString)
End Function



'''===================================================================================================================================================
'''<summary>
'''Removes all leading white-space characters from the specified string.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<returns>The string that remains after all white-space characters are removed from the start of the specified string. If no characters can be trimmed, the base string is unchanged.</returns>
'''===================================================================================================================================================
Public Function TrimLeft(ByVal baseString As String) As String
    TrimLeft = VBA.LTrim$(baseString)
End Function



'''===================================================================================================================================================
'''<summary>
'''Returns a new string truncated after the specified length with an extension added to note the truncation.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="count">Number of characters to truncate after.</param>
'''<param name="extension">String to append to the end of the truncated string to note the truncation</param>
'''<returns>A string truncated to count characters, with the added extension appended to the end.</returns>
'''<remarks>Implementation courtesy of Robert Todar, <see href="https://github.com/todar/VBA-Strings/blob/master/StringFunctions.bas#L185">VBA-Strings</see>.</remarks>
'''===================================================================================================================================================
Public Function Truncate(ByVal baseString As String, ByVal count As Long, Optional ByVal extension As String = "...") As String
    If Len(baseString) <= count Then
        Truncate = baseString
        Exit Function
    End If

    Truncate = Left(baseString, count - Len(extension)) & extension
End Function



'''===================================================================================================================================================
'''<summary>
'''Encode a string to a URL-friendly encoding.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="spaceAsPlus">Specify to encode a space as "+" instead of "%20"</param>
'''<returns>A string truncated to count characters, with the added extension appended to the end.</returns>
'''<remarks>Implementation is based on <see href="https://stackoverflow.com/a/218199/10802255">.</remarks>
'''===================================================================================================================================================
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



'''===================================================================================================================================================
'''<summary>
'''Wraps a paragraph of text so each line is at most lineWidth characters long, plus or minus a tolerance
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="lineWidth">Approximately the length of a line in number of characters.</param>
'''<param name="newLineChars">Characters to use as the new line separators. Default is `vbNewLine`.</param>
'''<param name="chopAtExactly">If specified, all traditional delimiters are ignored and each line will be exactly lineWidth characters wide. Default is FALSE.</param>
'''<param name="tolerance">Specifies the limit for how far away from the lineWidth to find the nearest wrap location. Default is 4 characters.</param>
'''<returns>Wrapped text delineated with a new line character</returns>
'''===================================================================================================================================================
Public Function Wrap(ByVal baseString As String, ByVal lineWidth As Long, Optional ByVal newLineChar As String = vbNewLine, Optional ByVal chopAtExactly As Boolean = False, _
    Optional ByVal tolerance As Long = 4) As String
    'TODO - allow for indents:
    'initial_indent: (default: '') String that will be prepended to the first line of wrapped output. Counts towards the length of the first line. The empty string is not indented.
    'subsequent_indent: (default: '') String that will be prepended to all lines of wrapped output except the first. Counts towards the length of each line except the first.
    If chopAtExactly Then
        Dim endPos As Long
        Do
            endPos = endPos + lineWidth
            Wrap = Wrap & Mid$(baseString, endPos - lineWidth + 1, lineWidth) & newLineChar
        Loop While endPos < Len(baseString)
        Wrap = Wrap & Mid$(baseString, endPos + 1)
        Exit Function
    
    Else
        Dim pos As Long
        Dim separators As Variant: separators = Array(" ", "-", "/", "\")
        Dim nextWrapLocn As Long
        Dim firstWrapLocn As Long
        Dim lastWrapLocn As Long
        Do
            firstWrapLocn = Strings.IndexOfAny(baseString, separators, nextWrapLocn + lineWidth - tolerance)
            lastWrapLocn = Strings.LastIndexOfAny(baseString, separators, nextWrapLocn + lineWidth + tolerance, tolerance * 2)
            nextWrapLocn = IIf(firstWrapLocn > lastWrapLocn, firstWrapLocn, lastWrapLocn)
'            If Mid$(baseString, nextWrapLocn, 1) = "\" Or Mid$(baseString, nextWrapLocn, 1) = "/" Then
'                nextWrapLocn = nextWrapLocn + 1 'Edge case to not split in between double slash escape sequences
'            End If
            
            Wrap = Wrap & VBA.Trim$(Mid$(baseString, pos + 1, nextWrapLocn - pos)) & newLineChar
            pos = nextWrapLocn
        Loop While nextWrapLocn <> -1 And nextWrapLocn + lineWidth + tolerance < Len(baseString)
        Wrap = Wrap & Mid$(baseString, pos + 1)
        Exit Function
    End If
End Function

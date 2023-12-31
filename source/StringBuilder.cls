VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBA-Strings")
'@IgnoreModule ProcedureNotUsed, ProceudreNotUsed
''===================================================================================================================================================
'' VBA-Strings
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
''    * This class represents a string-like object whose value is a mutable sequence of characters which offers superior performance when extensive
''      string manipulation operations are required.
''    * For usage instructions, including for when to use regular strings over the StringBuilder class, refer to the documentation at:
''      https://github.com/noah-severyn/VBA-Strings
''    * Based on VBA-StringBuffer (https://github.com/cristianbuse/VBA-StringBuffer), Copyright (c) 2019 Ion Cristian Buse, licensed under the MIT license.
''===================================================================================================================================================

Option Explicit

'''<summary>
'''Specifies the direction for processing the string.
'''</summary>
    '''<summary>
    '''Start processing at the start of the string and move forward.
    '''</summary>
    '''<summary>
    '''Start processing at the end of the string and move backwards.
    '''</summary>
Public Enum ProcessDirection
    FromStart
    FromEnd
End Enum

'Internal BUFFER Struct
Private Type Buffer
    text_ As String
    endIndex_ As Long
    capacity_ As Long
End Type

'Class members
Private pBuffer As Buffer
Private pMaxCapacity As Long


'''===================================================================================================================================================
'''<summary>
'''Gets the number of characters that can be contained in the memory allocated by the current instance.
'''</summary>
'''<returns>The maximum number of characters that can be contained in the memory allocated by the current instance. Its value can range from Length to MaxCapacity.</returns>
'''===================================================================================================================================================
Public Property Get CurrentCapacity() As Long
    CurrentCapacity = pBuffer.capacity_
End Property
'''===================================================================================================================================================
'''<summary>
'''Gets the length of the current StringBuilder object.
'''</summary>
'''<returns>The length of this instance.</returns>
'''===================================================================================================================================================
Public Property Get Length() As Long
    Length = pBuffer.endIndex_
End Property
'''===================================================================================================================================================
'''<summary>
'''Gets the character at the specified character position in this instance.
'''</summary>
'''<param name="index">The position of the character.</param>
'''<error cref="9">Index must be greater than zero.</error>
'''<returns>The Unicode character at position index.</returns>
'''=================================================================================================================================================== 'TODO add default value here
Public Property Get Chars(ByVal index As Long) As String
    If index < 0 Then
        Err.Raise 9, "StringBuilder.Chars", "Index must be greater than zero."
    End If

    If index > pBuffer.endIndex_ Then index = pBuffer.endIndex_
    Chars = Mid$(pBuffer.text_, index + 1, 1)
End Property
'Public Property Let Chars(ByVal index As Long) 'TODO allow to set Chars
'    mid#(pbuffer.text_,
'End Property
'''===================================================================================================================================================
'''<summary>
'''Sets the maximum capacity of this instance, regardless of the current memory allocation.
'''</summary>
'''<error cref="9">Maximum capacity must be greater than zero.</error>
'''===================================================================================================================================================
Public Property Let MaxCapacity(ByVal maximum As Long)
    If maximum <= 0 Then
        Err.Raise 9, "StringBuilder.MaxCapacity", "Maximum capacity must be greater than zero."
    End If
    pMaxCapacity = maximum
End Property
'''===================================================================================================================================================
'''<summary>
'''Gets the maximum capacity of this instance, regardless of the current memory allocation.
'''</summary>
'''<returns>The maximum number of characters this instance can hold.</returns>
'''<remarks>
'''The absolute maximum capacity is the max length allowed for variable length strings: 2,147,483,646
'''</remarks>
'''===================================================================================================================================================
Public Property Get MaxCapacity() As Long
    MaxCapacity = pMaxCapacity
End Property



'''===================================================================================================================================================
'''<summary>
'''Initializes a new instance of the StringBuilder class with the maximum capacity set to the max length allowed for variable length strings.
'''</summary>
'''===================================================================================================================================================
Private Sub Class_Initialize()
    pMaxCapacity = &H7FFFFFFF - 1 '2,147,483,646 (dec)
End Sub
'''===================================================================================================================================================
'''<summary>
'''Extend the buffer size if needed.
'''</summary>
'''<param name="minimumCapacity">Capacity to verify</param>
'''===================================================================================================================================================
Private Sub EnsureCapacity(ByVal minimumCapacity As Long)
    If minimumCapacity <= 0 Then
        Err.Raise 9, "StringBuilder.EnsureCapacity", "Minimum capacity must be greater than zero."
    End If
    If minimumCapacity > pBuffer.capacity_ Then
        Dim oldCapacity As Long: oldCapacity = pBuffer.capacity_
        
        If CDbl(minimumCapacity) * 2# > CDbl(&H7FFFFFFF - 1) Then
            pBuffer.capacity_ = &H7FFFFFFF - 1
        ElseIf CDbl(minimumCapacity) * 2# > CDbl(pMaxCapacity) Then
            pBuffer.capacity_ = pMaxCapacity
        Else
            pBuffer.capacity_ = minimumCapacity * 2
        End If

        pBuffer.text_ = pBuffer.text_ & VBA.space$(pBuffer.capacity_ - oldCapacity)
    End If
End Sub




'''===================================================================================================================================================
'''<summary>
'''Appends a copy of the specified string to this instance.
'''</summary>
'''<param name="text">The string to append.</param>
'''===================================================================================================================================================
Public Sub Append(ByRef text As String)
    Dim addedLength As Long: addedLength = VBA.Len(text)
    If addedLength = 0 Then Exit Sub
    Dim newEndIndex As Long: newEndIndex = pBuffer.endIndex_ + addedLength
    
    'EnsureCapacity already checks for: newEndIndex > m_buffer.capacity_
    'When many appends are done (ex. 1 character at a time for a million times) then the extra 'If' is faster than exiting the scope of this function
    'i.e. avoid pushing a new stack frame at the top of the call stack
    If newEndIndex > pBuffer.capacity_ Then EnsureCapacity newEndIndex

    'Replace unused characters with the new text after the last used position. The Mid Statement can be used on the left hand of the assignment to
    'in-place replace characters: https://en.wikibooks.org/wiki/Visual_Basic/Built_In_String_Functions#Indexing_and_Substrings
    Mid$(pBuffer.text_, pBuffer.endIndex_ + 1, addedLength) = text
    
    'Store last character position
    pBuffer.endIndex_ = newEndIndex
End Sub



'''===================================================================================================================================================
'''<summary>
'''Appends a copy of the specified strings to this instance.
'''</summary>
'''<param name="args()">One or more strings to append.</param>
'''<error cref="9">Argument at a position could not be converted to a string.</error>
'''===================================================================================================================================================
Public Sub AppendMultiple(ParamArray args() As Variant)
    Dim idx As Long
    Dim idxInner As Long
    On Error GoTo AppendInvalidArg
    For idx = LBound(args) To UBound(args)
        If IsArray(args(idx)) Then
            For idxInner = 0 To UBound(args(idx))
                Append CStr(args(idx)(idxInner))
            Next idxInner
        Else
            Append CStr(args(idx))
        End If
    Next idx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "StringBuilder.AppendJoin", "Argument at position " & idx * 2 + 1 & " could not be converted to a string."
End Sub



'''===================================================================================================================================================
'''<summary>
'''Concatenates the strings of the provided array, using the specified separator between each string, then appends the result to the current instance of the string builder.
'''</summary>
'''<param name="separator">Separator to append in between each value.</param>
'''<param name="args()">One or more strings to append.</param>
'''<error cref="9">Argument at a position could not be converted to a string.</error>
'''===================================================================================================================================================
Public Sub AppendJoin(ByVal separator As String, ByRef args() As String)
    Dim idx As Long
    Dim idxInner As Long
    On Error GoTo AppendInvalidArg
    For idx = 0 To UBound(args)
        If IsArray(args(idx)) Then
            For idxInner = 0 To UBound(args(idx))
                Append separator
                '@Ignore DefaultMemberRequired
                Append = Append & CStr(args(idx)(idxInner))
            Next idxInner
        Else
            Append separator
            Append CStr(args(idx))
        End If
    Next idx
    On Error GoTo 0
    Exit Function
    
AppendInvalidArg:
    Err.Raise 5, "StringBuilder.AppendJoin", "Argument at position " & idx * 2 + 1 & " could not be converted to a string."
End Sub



'''===================================================================================================================================================
'''<summary>
'''Appends the default line terminator, or a copy of a specified string and the default line terminator, to the end of this instance.
'''</summary>
'''<param name="text">The string to append. If this is omitted then only a new line is appended.</param>
'''===================================================================================================================================================
Public Sub AppendLine(Optional ByVal text As String)
    Append vbNewLine
    If text <> vbNullString Then
        Append text
    End If
End Sub



'''===================================================================================================================================================
'''<summary>
'''Removes all characters from the current instance.
'''</summary>
'''===================================================================================================================================================
Public Sub Clear()
    pBuffer.text_ = vbNullString
    pBuffer.endIndex_ = 0
    pBuffer.capacity_ = 0
End Sub


Public Function Contains(ByVal stringToFind As String, Optional ByVal compare As StringComparison = 0) As Boolean
    Contains = IndexOf(Me.ToString, stringToFind, , compare) >= 0
End Function


Public Function ContainsAfter(ByVal stringToFind As String, Optional ByVal startIndex As Long, Optional ByVal compare As StringComparison = 0) As Boolean
    ContainsAfter = IndexOf(Me.ToString, stringToFind, startIndex, compare) >= 0
End Function


Public Function ContainsBefore(ByVal stringToFind As String, Optional ByVal endIndex As Long, Optional ByVal compare As StringComparison = 0) As Boolean
    Dim newString As String: newString = Left(Me.ToString, endIndex)
    ContainsBefore = IndexOf(newString, stringToFind, , compare) >= 0
End Function


Public Function ContainsAny(ByVal stringsToFind As Variant, Optional ByVal startIndex As Long, Optional ByVal compare As StringComparison = 0) As Boolean
    Dim searchString As String: searchString = Me.ToString
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        ContainsAny = IndexOf(searchString, stringsToFind(idx), startIndex, compare) >= 0
        If ContainsAny Then Exit Function
    Next idx
End Function


Public Function CountSubstring(ByVal substringToFind As String, Optional ByVal compare As StringComparison = 0) As Long
    Dim baseString As String: baseString = Me.ToString
    Dim locn As Long: locn = IndexOf(baseString, substringToFind, locn)

    Do While locn >= 0
        locn = IndexOf(baseString, substringToFind, locn + 1, compare)
        CountSubstring = CountSubstring + 1
    Loop
End Function



'0-based index
Public Function IndexOf(ByVal searchString As String, Optional ByVal startPos As Long = 0, Optional ByVal compare As StringComparison = 0) As Long
    If startPos < 0 Then
        Err.Raise 9, "StringBuilder.IndexOf", "Start position must be greater than zero."
    ElseIf startPos > Len(pBuffer.endIndex_) Then
        Err.Raise 9, "StringBuilder.IndexOf", "Start position must be less than this instance's length."
    End If

    Dim tempString As String: tempString = Me.Substring(startPos, pBuffer.endIndex_ - startPos)
    If compare Mod 2 = 1 Then
        searchString = LCase$(searchString)
        tempString = LCase$(tempString)
    End If

    If compare = Binary Or compare = BinaryIgnoreCase Then
        IndexOf = InStr(startPos + 1, tempString, searchString, vbBinaryCompare) - 1
    ElseIf compare = text Or compare = TextIngnoreCase Then
        IndexOf = InStr(startPos + 1, tempString, searchString, vbTextCompare) - 1
    ElseIf compare = Database Or compare = DatabaseIgnoreCase Then
        IndexOf = InStr(startPos + 1, tempString, searchString, vbDatabaseCompare) - 1
    Else
        IndexOf = InStr(startPos + 1, tempString, searchString) - 1
    End If
    If IndexOf >= pBuffer.endIndex_ Then IndexOf = -1
End Function



'''===================================================================================================================================================
'''<summary>
'''Inserts one or more copies of a specified string into this instance at the specified character position.
'''</summary>
'''<param name="startIndex">The position in this instance where insertion begins.</param>
'''<param name="text">The string to insert.</param>
'''<param name="repetitions">The number of times to insert value.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Number of repetitions cannot be less than 0.</error>
'''===================================================================================================================================================
Public Sub Insert(ByVal startIndex As Long, ByVal text As String, Optional ByVal repetitions As Long = 1)
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, "StringBuilder.Insert", "Start index must be greater than zero."
    ElseIf repetitions < 1 Then
        Err.Raise 9, "StringBuilder.Insert", "Number of repetitions cannot be less than 0."
    End If
    
    If repetitions > 1 Then
        text = Strings.Replace(VBA.String(repetitions, " "), " ", text)
    End If
    
    Dim addedLength As Long: addedLength = VBA.Len(text)
    If addedLength = 0 Then Exit Sub 'Nothing to insert
    Dim newEndIndex As Long: newEndIndex = pBuffer.endIndex_ + addedLength
    Dim shiftLength As Long: shiftLength = pBuffer.endIndex_ - startIndex
    
    'EnsureCapacity already checks for: newEndIndex > m_buffer.capacity_
    'When many appends are done (ex. 1 character at a time for a million times) then the extra 'If' is faster than exiting the scope of this function
    'i.e. avoid pushing a new stack frame at the top of the call stack
    If newEndIndex > pBuffer.capacity_ Then EnsureCapacity newEndIndex
    
    'Shift Text Right
    Mid$(pBuffer.text_, startIndex + addedLength + 1, shiftLength) = VBA.Mid$(pBuffer.text_, startIndex + 1, shiftLength + 1)
    
    'Replace unused characters with the new text after the last used position. The Mid Statement can be used on the left hand of the assignment to
    'in-place replace characters: https://en.wikibooks.org/wiki/Visual_Basic/Built_In_String_Functions#Indexing_and_Substrings
    Mid$(pBuffer.text_, startIndex + 1, addedLength) = text
    
    'Update last character position
    pBuffer.endIndex_ = newEndIndex
End Sub



Public Sub Interpolate(ParamArray args() As Variant)
    Dim argCount As Long: argCount = UBound(args) - LBound(args) + 1
    If (argCount) Mod 2 <> 0 Then
        Err.Raise 5, "StringBuilder.Interpolate", "Invalid number of parameters. The interpolation parameters must be provided in pairs."
    End If
    
    Dim argIdx As Long
    On Error GoTo InterpolateInvalidArg
    For argIdx = 0 To argCount - 2 Step 2 '
        Me.Replace "{" & CStr(args(argIdx)) & "}", CStr(args(argIdx + 1))
    Next argIdx
    On Error GoTo 0
    Exit Sub
    
InterpolateInvalidArg:
    Err.Raise 5, "StringBuilder.Interpolate", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Sub


Public Sub Interpolate2(ByVal baseString As String, ByVal args As Dictionary)
    Dim argIdx As Long
    Dim key As Variant
    On Error GoTo Interpolate2InvalidArg
    For Each key In args.Keys()
        Interpolate2 = Strings.Replace(baseString, "{" & CStr(key) & "}", CStr(args.item(key)))
        argIdx = argIdx + 1
    Next key
    On Error GoTo 0
    Exit Sub
    
Interpolate2InvalidArg:
    Err.Raise 5, "StringBuilder.Interpolate2", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Sub


Public Function LastIndexOf(ByVal searchString As String, Optional ByVal startPos As Long = -2, Optional ByVal compare As StringComparison = 0) As Long
    Dim tempString As String: tempString = Me.Substring(startPos, pBuffer.endIndex_ - startPos)
    If compare Mod 2 = 1 Then
        searchString = LCase$(searchString)
        tempString = LCase$(tempString)
    End If
    Dim pos As Long
    If compare = Binary Or compare = BinaryIgnoreCase Then
        pos = InStrRev(tempString, searchString, startPos + 1, vbBinaryCompare) - 1
    ElseIf compare = text Or compare = TextIngnoreCase Then
        pos = InStrRev(tempString, searchString, startPos + 1, vbTextCompare) - 1
    ElseIf compare = Database Or compare = DatabaseIgnoreCase Then
        pos = InStrRev(tempString, searchString, startPos + 1, vbDatabaseCompare) - 1
    Else
        pos = InStrRev(tempString, searchString, startPos + 1) - 1
    End If
    
    If pos >= pBuffer.endIndex_ Then pos = LastIndexOf(searchString, pos - 1, compare)
    LastIndexOf = pos
End Function



'''===================================================================================================================================================
'''<summary>
'''Removes the specified range of characters from this instance.
'''</summary>
'''<param name="startIndex">The zero-based position in this instance where removal begins.</param>
'''<param name="count">TThe number of characters to remove.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="5">Count must be greater than zero.</error>
'''===================================================================================================================================================
Public Sub Remove(ByVal startIndex As Long, ByVal count As Long)
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, "StringBuilder.Delete", "Start index must be greater than zero."
    ElseIf count < 0 Then
        Err.Raise 5, "StringBuilder.Delete", "Count must be greater than zero."
    ElseIf count = 0 Then
        Exit Sub 'Nothing to delete
    End If
    
    'Check if a simple shift of the endIndex would suffice
    If startIndex + count > pBuffer.endIndex_ Then
        'Ignoring characters that were marked for deletion
        pBuffer.endIndex_ = startIndex - 1
        Exit Sub
    End If
    
    Dim shiftLength As Long
    shiftLength = pBuffer.endIndex_ - startIndex - count
    
    'Shift Text Left
    Mid$(pBuffer.text_, startIndex + 1, shiftLength) = VBA.Mid$(pBuffer.text_, startIndex + count + 1, shiftLength)
    
    'Update last character position
    pBuffer.endIndex_ = pBuffer.endIndex_ - count
End Sub



'''===================================================================================================================================================
'''<summary>
'''Removes the specified number of characters from the end this instance.
'''</summary>
'''<param name="count">The number of characters to remove.</param>
'''<error cref="5">Count must be greater than zero.</error>
'''<error cref="5">Count cannot be greater than the current length of this instance.</error>
'''===================================================================================================================================================
Public Sub RemoveFromEnd(ByVal count As Long)
    If count < 0 Then
        Err.Raise 5, "StringBuilder.RemoveFromEnd", "Count must be greater than zero."
    ElseIf count > pBuffer.endIndex_ Then
        Err.Raise 5, "StringBuilder.RemoveFromEnd", "Count cannot be greater than the current length of this instance."
    ElseIf count = 0 Then
        Exit Sub 'Nothing to delete
    End If
    
    pBuffer.endIndex_ = pBuffer.endIndex_ - count
End Sub



'''===================================================================================================================================================
'''<summary>
'''Removes all instances of the specified characters from the end this instance until no more are remaining.
'''</summary>
'''<param name="stringToRemove">The text to remove.</param>
'''<error cref="9">String to remove must have length greater than 0 and cannot be null.</error>
'''===================================================================================================================================================
Public Sub RemoveFromEndWhile(ByVal stringToRemove As String)
    If stringToRemove = vbNullString Or Len(stringToRemove) = 0 Then
        Err.Raise 9, "StringBuilder.RemoveFromEndWhile", "String to remove must have length greater than 0 and cannot be null."
    End If
    
    Dim charCount As Long: charCount = Len(stringToRemove)
    
    Do While Me.Substring(pBuffer.endIndex_ - charCount, charCount) = stringToRemove
        pBuffer.endIndex_ = pBuffer.endIndex_ - charCount
    Loop
End Sub



Public Sub RemoveNonPrintableChars()
    'Does not remove new line characters
    Dim idx As Long
    Dim currCharCode As Long
    For idx = 0 To pBuffer.endIndex_
        currCharCode = AscW(Mid$(pBuffer.text_, idx + 1, 1)) And &HFFFF& 'Equivalent to Strings.AscW2
        If currCharCode > 31 Or currCharCode = 13 Or currCharCode = 10 Then
            Me.Remove idx, 1
        End If
    Next idx
End Sub



'''===================================================================================================================================================
'''<summary>
'''Overwrites the specified range of characters from this instance with the specified replacement text.
'''</summary>
'''<param name="startIndex">The zero-based position in this instance where overwriting begins.</param>
'''<param name="count">The number of characters to remove.</param>
'''<param name="replacementText">The number of characters to remove.</param>
'''===================================================================================================================================================
Public Sub Overwrite(ByVal startIndex As Long, ByVal count As Long, ByVal replacementText As String)
    'Validate Input
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, "StringBuilder.Overwrite", "Invalid startIndex"
    ElseIf count < 0 Then
        Err.Raise 5, "StringBuilder.Overwrite", "Invalid count"
    ElseIf count = 0 Then
        Exit Sub 'Nothing to replace
    End If
    
    Dim usedLength As Long
    
    'Compute usable length
    If startIndex + count > pBuffer.endIndex_ + 1 Then
        usedLength = pBuffer.endIndex_ - startIndex
    Else
        usedLength = count
    End If
    
    Mid$(pBuffer.text_, startIndex + 1, usedLength) = replacementText
End Sub



'''===================================================================================================================================================
'''<summary>
'''Replaces all occurrences of a specified string in this instance with another specified string.
'''</summary>
'''<param name="oldString">The string to replace.</param>
'''<param name="newString">The string that replaces oldValue.</param>
'''<param name="numberOfSubstitutions">The number of substitutions to make.</param>
'''<param name="direction">The direction to start substituting from.</param>
'''===================================================================================================================================================
Public Sub Replace(ByVal oldString As String, ByVal newString As String, Optional ByVal numberOfSubstitutions As Long = -1, Optional ByVal direction As ProcessDirection = ProcessDirection.FromStart)
    Me.ReplaceBetween oldString, newString, 0, pBuffer.endIndex_, numberOfSubstitutions, direction
End Sub



'''===================================================================================================================================================
'''<summary>
'''Replaces, within a substring of this instance, all occurrences of a specified string with another specified string.
'''</summary>
'''<param name="oldString">The string to replace.</param>
'''<param name="newString">The string that replaces oldValue.</param>
'''<param name="startIndex">The position in this instance where the substring begins.</param>
'''<param name="count">The length of the substring.</param>
'''<param name="numberOfSubstitutions">The number of substitutions to make.</param>
'''<param name="direction">The direction to start substituting from.</param>
'''===================================================================================================================================================
Public Sub ReplaceBetween(ByVal oldString As String, ByVal newString As String, ByVal startIndex As Long, ByVal count As Long, Optional ByVal numberOfSubstitutions As Long = -1, Optional ByVal direction As ProcessDirection = ProcessDirection.FromStart)
    If oldString = vbNullString Then
        Err.Raise 5, "StringBuilder.Replace", "String to find cannot be null"
    ElseIf startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, "StringBuilder.Replace", "Invalid startIndex"
    ElseIf count < -1 Then
        Err.Raise 5, "StringBuilder.Replace", "Invalid length"
    ElseIf count = 0 Or numberOfSubstitutions = 0 Then
        Exit Sub 'Nothing to replace
    End If
    
    Dim pos As Long: pos = startIndex
    Dim countSubs As Long
    Dim maxSubs As Long
    If numberOfSubstitutions = -1 Then
        maxSubs = 2147483647
    Else
        maxSubs = numberOfSubstitutions
    End If
    
    If direction = FromStart Then
        Do While pos < count And countSubs < maxSubs
            pos = Me.IndexOf(oldString, pos)
            If pos = -1 Then Exit Sub 'No more replacements to make
            Me.Remove pos, Len(oldString)
            Me.Insert pos, newString, 1
            countSubs = countSubs + 1
        Loop
    Else
        Do While pos > pBuffer.endIndex_ - count And countSubs < maxSubs
            pos = Me.LastIndexOf(oldString, pos)
            If pos = -1 Then Exit Sub 'No more replacements to make
            Me.Remove pos, Len(oldString)
            Me.Insert pos, newString, 1
            countSubs = countSubs + 1
        Loop
    End If
End Sub





Public Sub ReplaceNewLineChars(Optional ByVal replacement As String = " ")
    Me.Replace vbCrLf, replacement
    Me.Replace vbCr, replacement
    Me.Replace vbLf, replacement
End Sub



Public Sub ReplaceNonBreakingSpaces()
    Me.Replace Chr$(160), " "
End Sub


'''===================================================================================================================================================
'''<summary>
'''Returns a string in which the character order of this instance is reversed.
'''</summary>
'''<returns>A reversed string.</returns>
'''===================================================================================================================================================
Public Sub Reverse()
    If pBuffer.endIndex_ > 0 Then
        Me.Overwrite 1, pBuffer.endIndex_, VBA.StrReverse(Me.Value)
    End If
End Sub



'''===================================================================================================================================================
'''<summary>
'''Retrieves a substring from this instance. The substring starts at a specified character position and continues to the end of the string or to the specified length.
'''</summary>
'''<param name="startIndex">The zero-based starting character position of a substring in this instance.</param>
'''<param name="count">The number of characters in the substring.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Count must be greater than zero.</error>
'''<returns>A string that is equivalent to the substring that begins at startIndex of baseString, the baseString if count is 0, or vbnullstring if startIndex is equal to the length of this instance.</returns>
'''===================================================================================================================================================
Public Function Substring(ByVal startIndex As Long, Optional ByVal count As Long = -1) As String
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, "StringBuilder.ToString", "Start index must be greater than zero."
    ElseIf count < -1 Then
        Err.Raise 9, "StringBuilder.ToString", "Count of chars must be greater than zero."
    ElseIf count = 0 Then
        Exit Function
    End If
    
    If count = -1 Then count = pBuffer.endIndex_ - startIndex
    Dim usedLength As Long
    
    'Compute usable length
    If startIndex + Length > pBuffer.endIndex_ + 1 Then
        usedLength = pBuffer.endIndex_ + 1 - startIndex
    Else
        usedLength = Length
    End If
    
    Substring = VBA.Mid$(pBuffer.text_, startIndex + 1, usedLength)
End Function


'''===================================================================================================================================================
'''<summary>
'''Retrieves a substring from this instance. The substring starts at the end of firstString and ends at the start location of secondString, optionally beginning the search for firstString as startIndex.
'''</summary>
'''<param name="baseString">Any valid string.</param>
'''<param name="firstString">String marking the start of the substring.</param>
'''<param name="secondString">String marking the end of the substring</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Value of first string was not found in the base string.</error>
'''<error cref="9">Value of second string was not found in the base string.</error>
'''<param name="startIndex">The zero-based starting character position to begin searching for firstString at.</param>
'''==================================================================================================================================================
Public Function SubstringBetween(ByVal firstString As String, ByVal secondString As String, Optional ByVal startIndex As Long = 0) As String
    If startIndex < 0 Then
        Err.Raise 9, "StringBuilder.SubstringBetween", "Start index must be greater than zero."
    End If
    Dim startPos As Long: startPos = Me.IndexOf(firstString, startIndex) + Len(firstString)
    If startPos = -1 Then
        Err.Raise 9, "StringBuilder.SubstringBetween", "Value of first string was not found in the base string."
    End If
    Dim endPos As Long: endPos = Me.IndexOf(secondString, startPos)
    If endPos = -1 Then
        Err.Raise 9, "Strings.SubstringBetween", "Value of second string was not found in the base string."
    End If
    
    
    
    SubstringBetween = Me.Substring(startPos, endPos - startPos)
End Function



'''===================================================================================================================================================
'''<summary>
'''Converts the value of this instance to a String.
'''</summary>
'''<param name="startIndex">The starting position of the substring in this instance.</param>
'''<param name="count">The length of the substring. If count + startIndex is longer than the length of this instance then the count is adjusted to the maximum allowed length.</param>
'''<error cref="9">Start index must be greater than zero.</error>
'''<error cref="9">Count of chars must be greater than zero.</error>
'''<returns>A string whose value is the same as the specified substring of this instance.</returns>
'''<remarks>You must call the ToString method to convert the StringBuilder object to a String object before you can pass the string represented by the
'''StringBuilder object to a method that has a String parameter or display it in the user interface.</remarks>
'''===================================================================================================================================================
Public Function ToString(Optional ByVal startIndex As Long = 0, Optional ByVal count As Long = 0) As String
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, "StringBuilder.ToString", "Start index must be greater than zero."
    ElseIf count < 0 Then
        Err.Raise 9, "StringBuilder.ToString", "Count of chars must be greater than zero."
    ElseIf startIndex + count < pBuffer.endIndex_ Then
        count = pBuffer.endIndex_ - startIndex
    ElseIf pBuffer.endIndex_ = 0 Then
        ToString = vbNullString
        Exit Function
    End If
    
    If startIndex > 0 Then
        ToString = Me.Substring(startIndex, count)
    Else
        ToString = VBA.Left$(pBuffer.text_, pBuffer.endIndex_)
    End If
End Function


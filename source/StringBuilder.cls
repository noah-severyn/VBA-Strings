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


'===============================================================================
'Returns the capacity of the string i.e. total length of buffer
'===============================================================================
Public Property Get CurrentCapacity() As Long
    CurrentCapacity = pBuffer.capacity_
End Property

'===============================================================================
'Returns the length of the string i.e. total number of used characters
'===============================================================================
Public Property Get Length() As Long
    Length = pBuffer.endIndex_
End Property


Public Property Get Char(ByVal index As Long) As String
    Char = Mid$(pBuffer.text_, index + 1, 1)
End Property


'===============================================================================
'Returns the maximum capacity of the string i.e. maximum the buffer will be allowed to be set
'===============================================================================
Public Property Let MaxCapacity(ByVal maximum As Long)
    pMaxCapacity = maximum
End Property
Public Property Get MaxCapacity() As Long
    MaxCapacity = pMaxCapacity
End Property


Private Sub Class_Initialize()
    'Maximum string length allowed by VBA for a variable-length string
    pMaxCapacity = &H7FFFFFFF - 1 '2,147,483,646 (dec)
End Sub

'****************************************************************************************************************************************************
'Extend buffer size if needed
'****************************************************************************************************************************************************
Private Sub EnsureCapacity(ByVal minimumCapacity As Long)
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



'****************************************************************************************************************************************************
'Append new text to buffer
'****************************************************************************************************************************************************
Public Sub Append(ByRef textToAppend As String)
    Dim addedLength As Long: addedLength = VBA.Len(textToAppend)
    If addedLength = 0 Then Exit Sub
    Dim newEndIndex As Long: newEndIndex = pBuffer.endIndex_ + addedLength
    
    'EnsureCapacity already checks for: newEndIndex > m_buffer.capacity_
    'When many appends are done (ex. 1 character at a time for a million times) then the extra 'If' is faster than exiting the scope of this function
    'i.e. avoid pushing a new stack frame at the top of the call stack
    If newEndIndex > pBuffer.capacity_ Then EnsureCapacity newEndIndex

    'Replace unused characters with the new text after the last used position. The Mid Statement can be used on the left hand of the assignment to
    'in-place replace characters: https://en.wikibooks.org/wiki/Visual_Basic/Built_In_String_Functions#Indexing_and_Substrings
    Mid$(pBuffer.text_, pBuffer.endIndex_ + 1, addedLength) = textToAppend
    
    'Store last character position
    pBuffer.endIndex_ = newEndIndex
End Sub

Public Sub AppendMultiple(ParamArray values() As Variant)
    Dim idx As Long
    For idx = LBound(values) To UBound(values)
        Append CStr(values(idx))
    Next idx
End Sub

Public Sub AppendJoin(ByVal separator As String, ByRef values() As String)
    Dim idx As Long
    For idx = LBound(values) To UBound(values)
        Append separator
        Append CStr(values(idx))
    Next idx
End Sub

Public Sub AppendLine(Optional ByVal textToAppend As String)
    Append Chr$(10)
    Append textToAppend
End Sub




'****************************************************************************************************************************************************
'Reset buffer members for (re)usage
'****************************************************************************************************************************************************
Public Sub Clear()
    pBuffer.text_ = vbNullString
    pBuffer.endIndex_ = 0
    pBuffer.capacity_ = 0
End Sub


Public Function Contains(ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    Contains = IndexOf(Me.ToString, stringToFind, , compare) >= 0
End Function


Public Function ContainsAfter(ByVal stringToFind As String, Optional ByVal startIndex As Long, Optional ByVal compare As Comparison = 0) As Boolean
    ContainsAfter = IndexOf(Me.ToString, stringToFind, startIndex, compare) >= 0
End Function


Public Function ContainsBefore(ByVal stringToFind As String, Optional ByVal endIndex As Long, Optional ByVal compare As Comparison = 0) As Boolean
    Dim newString As String: newString = Left(Me.ToString, endIndex)
    ContainsBefore = IndexOf(newString, stringToFind, , compare) >= 0
End Function


Public Function ContainsAny(ByVal stringsToFind As Variant, Optional ByVal startIndex As Long, Optional ByVal compare As Comparison = 0) As Boolean
    Dim searchString As String: searchString = Me.ToString
    Dim idx As Long
    For idx = 0 To UBound(stringsToFind)
        ContainsAny = IndexOf(searchString, stringsToFind(idx), startIndex, compare) >= 0
        If ContainsAny Then Exit Function
    Next idx
End Function


Public Function CountSubstring(ByVal substringToFind As String, Optional ByVal compare As Comparison = 0) As Long
    Dim baseString As String: baseString = Me.ToString
    Dim locn As Long: locn = IndexOf(baseString, substringToFind, locn)

    Do While locn >= 0
        locn = IndexOf(baseString, substringToFind, locn + 1, compare)
        CountSubstring = CountSubstring + 1
    Loop
End Function



'0-based index
Public Function IndexOf(ByVal stringToFind As String, Optional ByVal startPos As Long = 0, Optional ByVal compare As Comparison = 0) As Long
    'TODO - .IndexOf I dont't think the case comparison would work
    If compare Mod 2 = 1 Then
        'searchString = LCase$(searchString)
        stringToFind = LCase$(stringToFind)
    End If

    If compare = Binary Or BinaryIgnoreCase Then
        IndexOf = InStr(startPos + 1, pBuffer.text_, stringToFind, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        IndexOf = InStr(startPos + 1, pBuffer.text_, stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        IndexOf = InStr(startPos + 1, pBuffer.text_, stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOf = InStr(startPos + 1, pBuffer.text_, stringToFind) - 1
    End If
    If IndexOf >= pBuffer.endIndex_ Then IndexOf = -1
End Function



'****************************************************************************************************************************************************
'Insert new text into buffer
'****************************************************************************************************************************************************
Public Sub Insert(ByVal startIndex As Long, ByRef textToInsert As String, Optional ByVal repetitions As Long = 1)
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Insert", "Invalid startIndex"
    ElseIf repetitions < 1 Then
        Err.Raise 9, TypeName(Me) & ".Insert", "Number of repetitions cannot be less than 0"
    End If
    
    If repetitions > 1 Then
        textToInsert = strings.Replace(VBA.String(repetitions, " "), " ", textToInsert)
    End If
    
    Dim addedLength As Long: addedLength = VBA.Len(textToInsert)
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
    Mid$(pBuffer.text_, startIndex + 1, addedLength) = textToInsert
    
    'Update last character position
    pBuffer.endIndex_ = newEndIndex
End Sub



Public Sub Interpolate(ParamArray args() As Variant)
    Dim argCount As Long: argCount = UBound(args) - LBound(args) + 1
    If (argCount) Mod 2 <> 0 Then
        Err.Raise 5, TypeName(Me) & ".Interpolate", "Invalid number of parameters. The interpolation parameters must be provided in pairs."
    End If
    
    Dim argIdx As Long
    On Error GoTo InterpolateInvalidArg
    For argIdx = 0 To argCount - 2 Step 2 '
        Me.Replace "{" & CStr(args(argIdx)) & "}", CStr(args(argIdx + 1))
    Next argIdx
    On Error GoTo 0
    Exit Sub
    
InterpolateInvalidArg:
    Err.Raise 5, TypeName(Me) & ".Interpolate", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Sub


Public Function Interpolate2(ByVal baseString As String, ByVal args As Dictionary)
    Dim argIdx As Long
    Dim key As Variant
    On Error GoTo Interpolate2InvalidArg
    For Each key In args.Keys()
        Me.Replace "{" & CStr(key) & "}", CStr(args(key))
    Next key
    On Error GoTo 0
    Exit Function
    
Interpolate2InvalidArg:
    Err.Raise 5, TypeName(Me) & ".Interpolate2", "Argument at position " & argIdx * 2 + 1 & " could not be converted to a string."
End Function


Public Function LastIndexOf(ByVal stringToFind As String, Optional ByVal startPos As Long = -2, Optional ByVal compare As Comparison = 0) As Long
    'TODO - .LastIndexOf how to handle case comparisons
    If compare Mod 2 = 1 Then
        'searchString = LCase$(searchString)
        stringToFind = LCase$(stringToFind)
    End If
    Dim pos As Long
    If compare = Binary Or BinaryIgnoreCase Then
        pos = InStrRev(pBuffer.text_, stringToFind, startPos + 1, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        pos = InStrRev(pBuffer.text_, stringToFind, startPos + 1, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        pos = InStrRev(pBuffer.text_, stringToFind, startPos + 1, vbDatabaseCompare) - 1
    Else
        pos = InStrRev(pBuffer.text_, stringToFind, startPos + 1) - 1
    End If
    
    If pos >= pBuffer.endIndex_ Then pos = LastIndexOf(stringToFind, pos - 1, compare)
    LastIndexOf = pos
End Function




'****************************************************************************************************************************************************
'Delete a portion of the text
'****************************************************************************************************************************************************
Public Sub Remove(ByVal startIndex As Long, ByVal count As Long)
    'Validate Input
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Delete", "Invalid startIndex"
    ElseIf count < 0 Then
        Err.Raise 5, TypeName(Me) & ".Delete", "Invalid count"
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


Public Sub RemoveFromEnd(ByVal count As Long)
    If count < 0 Or count > pBuffer.endIndex_ Then
        Err.Raise 5, TypeName(Me) & ".RemoveFromEnd", "Invalid count"
    ElseIf count = 0 Then
        Exit Sub 'Nothing to delete
    End If
    
    pBuffer.endIndex_ = pBuffer.endIndex_ - count
End Sub


Public Sub RemoveFromEndWhile(ByVal stringToRemove As String)
    If stringToRemove = vbNullString Or Len(stringToRemove) = 0 Then
        Err.Raise 9, TypeName(Me) & ".RemoveFromEndWhile", "String to remove must have length greater than 0 and cannot be null"
    End If
    
    Dim charCount As Long: charCount = Len(stringToRemove)
    
    Do While Me.Substring(pBuffer.endIndex_ - charCount, charCount) = stringToRemove
        pBuffer.endIndex_ = pBuffer.endIndex_ - charCount
    Loop
End Sub



Public Sub RemoveNonPrintableChars()
    'Does not remove new line characters
    Dim idx As Long
    Dim c As Long
    Dim currentCharCode As Long

    RemoveNonPrintableChars = String$(Len(baseString), Chr$(0))
    For idx = 0 To pBuffer.endIndex_
        currentCharCode = AscW(Mid$(pBuffer.text_, idx + 1, 1)) And &HFFFF& 'Strings.AscW2
        If currentCharCode > 31 Or currentCharCode = 13 Or currentCharCode = 10 Then
            Me.Remove idx, 1
        End If
    Next idx
End Sub




'****************************************************************************************************************************************************
'Replace a portion of the buffer with a given text
'****************************************************************************************************************************************************
Public Sub Overwrite(ByVal startIndex As Long, ByVal Length As Long, ByVal replacementText As String) 'TODO - .Overwrite redesign to be 0-based
    'Validate Input
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Overwrite", "Invalid startIndex"
    ElseIf Length < 0 Then
        Err.Raise 5, TypeName(Me) & ".Overwrite", "Invalid length"
    ElseIf Length = 0 Then
        Exit Sub 'Nothing to replace
    End If
    
    Dim usedLength As Long
    
    'Compute usable length
    If startIndex + Length > pBuffer.endIndex_ + 1 Then
        usedLength = pBuffer.endIndex_ - startIndex
    Else
        usedLength = Length
    End If
    
    Mid$(pBuffer.text_, startIndex + 1, usedLength) = replacementText
End Sub


Public Sub Replace(ByVal oldString As String, ByVal newString As String, Optional ByVal numberOfSubstitutions As Long = -1, Optional ByVal direction As ProcessDirection = ProcessDirection.FromStart)
    'TODO - .Replace implement direction
    Me.ReplaceBetween oldString, newString, 0, pBuffer.endIndex_, numberOfSubstitutions
End Sub


Public Sub ReplaceBetween(ByVal oldString As String, ByVal newString As String, ByVal startIndex As Long, ByVal Length As Long, Optional ByVal numberOfSubstitutions As Long = -1, Optional ByVal direction As ProcessDirection = ProcessDirection.FromStart)
'TODO - .ReplaceBetween implement direction
    'Validate Input
    If oldString = vbNullString Then
        Err.Raise 5, TypeName(Me) & ".Replace", "String to find cannot be null"
    ElseIf startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Replace", "Invalid startIndex"
    ElseIf Length < -1 Then
        Err.Raise 5, TypeName(Me) & ".Replace", "Invalid length"
    ElseIf Length = 0 Or numberOfSubstitutions = 0 Then
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
    
    Do While pos < pBuffer.endIndex_ And countSubs < maxSubs
        pos = Me.IndexOf(oldString, pos)
        If pos = -1 Then Exit Sub 'No more replacements to make
        Me.Remove pos, Len(oldString)
        Me.Insert pos, newString, 1
        countSubs = countSubs + 1
    Loop
End Sub





Public Sub ReplaceNewLineChars(Optional ByVal replacement As String = " ")
    Me.Replace vbCrLf, replacement
    Me.Replace vbCr, replacement
    Me.Replace vbLf, replacement
End Sub



Public Sub ReplaceNonBreakingSpaces()
    Me.Replace Chr$(160), " "
End Sub


'****************************************************************************************************************************************************
'Reverses the contained string
'****************************************************************************************************************************************************
Public Sub Reverse()
    If pBuffer.endIndex_ > 0 Then
        Me.Overwrite 1, pBuffer.endIndex_, VBA.StrReverse(Me.Value)
    End If
End Sub




'****************************************************************************************************************************************************
'Returns a substring
'****************************************************************************************************************************************************
Public Function Substring(ByVal startIndex As Long, ByVal Length As Long) As String
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Substring", "Invalid startIndex"
    ElseIf Length < 0 Then
        Err.Raise 5, TypeName(Me) & ".Substring", "Invalid length"
    ElseIf Length = 0 Then
        Exit Function
    End If
    
    Dim usedLength As Long
    
    'Compute usable length
    If startIndex + Length > pBuffer.endIndex_ + 1 Then
        usedLength = pBuffer.endIndex_ + 1 - startIndex
    Else
        usedLength = Length
    End If
    
    Substring = VBA.Mid$(pBuffer.text_, startIndex + 1, usedLength)
End Function



Public Function SubstringBetween(ByVal firstString As String, ByVal secondString As String, Optional ByVal startIndex As Long = 0) As String
    Dim startPos As Long: startPos = Me.IndexOf(firstString, startIndex) + Len(firstString)
    Dim endPos As Long: endPos = Me.IndexOf(secondString, startPos)
    SubstringBetween = Me.Substring(startPos, endPos - startPos)
End Function



Public Function ToString(Optional ByVal startIndex As Long = 0, Optional ByVal Length As Long = 0) As String
    'Validate Input
    If startIndex < 0 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".ToString", "Invalid startIndex"
    ElseIf Length < 0 Then
        Err.Raise 5, TypeName(Me) & ".ToString", "Invalid length"
    ElseIf pBuffer.endIndex_ = 0 Then
        ToString = vbNullString
        Exit Function
    End If
    
    If startIndex > 0 Then
        ToString = Me.Substring(startIndex, Length)
    Else
        ToString = VBA.Left$(pBuffer.text_, pBuffer.endIndex_)
    End If
End Function


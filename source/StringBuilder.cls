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
Option Explicit

'Based on VBA-StringBuffer, Copyright (c) 2019 Ion Cristian Buse, licensed under the MIT license.
'https://github.com/cristianbuse/VBA-StringBuffer
''=============================================================================
'' VBA StringBuffer
''-------------------------------------------------
'' https://github.com/cristianbuse/VBA-StringBuffer
''-------------------------------------------------
''
'' Copyright (c) 2019 Ion Cristian Buse
''
'' Permission is hereby granted, free of charge, to any person obtaining a copy
'' of this software and associated documentation files (the "Software"), to deal
'' in the Software without restriction, including without limitation the rights
'' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'' copies of the Software, and to permit persons to whom the Software is
'' furnished to do so, subject to the following conditions:
''
'' The above copyright notice and this permission notice shall be included in all
'' copies or substantial portions of the Software.
''
'' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'' SOFTWARE.
''=============================================================================
''==============================================================================
'' Description:
''    * A 'Java-like' string buffer
'' The main methods are (both using the 'Mid' statement for speed):
''    * Append (much faster than using regular VBA concatenation - & operator)
''    * Insert
'' Other useful methods:
''    * Delete
''    * Replace
''    * Reset
''    * Reverse
''    * Substring
'' Retrieve value by calling property:
''    * Value (default class member - can be omitted)
'' Notes:
''    * The buffer's capacity should only be increased by using the
''      EnsureCapacity method
''==============================================================================

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
Public Property Get capacity() As Long
    capacity = pBuffer.capacity_
End Property

'===============================================================================
'Returns the length of the string i.e. total number of used characters
'===============================================================================
Public Property Get length() As Long
    length = pBuffer.endIndex_
End Property

'===============================================================================
'Get the Used String
'Default class member. 'strBuffer.Value' can be also called as 'strBuffer'
'Open class in a text editor to see: Attribute [procName].VB_UserMemId = 0
'===============================================================================
Public Property Get Value() As String
    If pBuffer.endIndex_ > 0 Then
        Value = VBA.Left$(pBuffer.text_, pBuffer.endIndex_)
    End If
End Property

Public Property Get Char(ByVal index As Long) As String
    Char = Mid$(pBuffer.text_, index, 1)
End Property


Public Property Let MaxCapacity(ByVal capacity As Long)
    pMaxCapacity = capacity
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

        pBuffer.text_ = pBuffer.text_ & VBA.Space$(pBuffer.capacity_ - oldCapacity)
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

Public Sub AppendJoin(ByVal separator As String, values() As String)
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




'****************************************************************************************************************************************************
'Insert new text into buffer
'****************************************************************************************************************************************************
Public Sub Insert(ByVal startIndex As Long, ByRef textToInsert As String, Optional ByVal repetitions As Long = 1)
    If startIndex < 1 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Insert", "Invalid startIndex"
    End If
    
    Dim addedLength As Long: addedLength = VBA.Len(textToInsert)
    If addedLength = 0 Then Exit Sub 'Nothing to insert
    Dim newEndIndex As Long: newEndIndex = pBuffer.endIndex_ + addedLength
    Dim shiftLength As Long: shiftLength = pBuffer.endIndex_ - startIndex + 1
    
    'EnsureCapacity already checks for: newEndIndex > m_buffer.capacity_
    'When many appends are done (ex. 1 character at a time for a million times) then the extra 'If' is faster than exiting the scope of this function
    'i.e. avoid pushing a new stack frame at the top of the call stack
    If newEndIndex > pBuffer.capacity_ Then EnsureCapacity newEndIndex
    
    'Shift Text Right
    Mid$(pBuffer.text_, startIndex + addedLength, shiftLength) = VBA.Mid$(pBuffer.text_, startIndex, shiftLength)
    
    'Replace unused characters with the new text after the last used position. The Mid Statement can be used on the left hand of the assignment to
    'in-place replace characters: https://en.wikibooks.org/wiki/Visual_Basic/Built_In_String_Functions#Indexing_and_Substrings
    Mid$(pBuffer.text_, startIndex, addedLength) = textToInsert
    
    'Update last character position
    pBuffer.endIndex_ = newEndIndex
End Sub





'****************************************************************************************************************************************************
'Delete a portion of the text
'****************************************************************************************************************************************************
Public Sub Remove(ByVal startIndex As Long, ByVal count As Long)
    'Validate Input
    If startIndex < 1 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Delete", "Invalid startIndex"
    ElseIf count < 0 Then
        Err.Raise 5, TypeName(Me) & ".Delete", "Invalid length_"
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
    
    shiftLength = pBuffer.endIndex_ - startIndex - count + 1
    
    'Shift Text Left
    Mid$(pBuffer.text_, startIndex, shiftLength) _
        = VBA.Mid$(pBuffer.text_, startIndex + count, shiftLength)
    
    'Update last character position
    pBuffer.endIndex_ = pBuffer.endIndex_ - count
End Sub




'****************************************************************************************************************************************************
'Replace a portion of the buffer with a given text
'****************************************************************************************************************************************************
Public Sub Overwrite(ByVal startIndex As Long, ByVal length As Long, ByRef replacementText As String)
    'Validate Input
    If startIndex < 1 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Replace", "Invalid startIndex"
    ElseIf length < 0 Then
        Err.Raise 5, TypeName(Me) & ".Replace", "Invalid length"
    ElseIf length = 0 Then
        Exit Sub 'Nothing to replace
    End If
    
    Dim usedLength As Long
    
    'Compute usable length
    If startIndex + length_ > pBuffer.endIndex_ + 1 Then
        usedLength = pBuffer.endIndex_ + 1 - startIndex
    Else
        usedLength = length
    End If
    
    'Replace characters with the new text. The Mid Statement can be used on the left hand of the assignment to
    'in-place replace characters: https://en.wikibooks.org/wiki/Visual_Basic/Built_In_String_Functions#Indexing_and_Substrings
    Mid$(pBuffer.text_, startIndex, usedLength) = replacementText
End Sub


Public Function ToString() As String
    ToString = pBuffer.text_
End Function




'****************************************************************************************************************************************************
'Reverses the contained string
'****************************************************************************************************************************************************
Public Sub Reverse()
    If pBuffer.endIndex_ > 0 Then
        Me.Replace 1, pBuffer.endIndex_, VBA.StrReverse(Me.Value)
    End If
End Sub




'****************************************************************************************************************************************************
'Returns a substring
'****************************************************************************************************************************************************
Public Function Substring(ByVal startIndex As Long, ByVal length_ As Long) As String
    'Validate Input
    If startIndex < 1 Or startIndex > pBuffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Substring", "Invalid startIndex"
    ElseIf length_ < 0 Then
        Err.Raise 5, TypeName(Me) & ".Substring", "Invalid length_"
    ElseIf length_ = 0 Then
        Exit Function
    End If
    '
    Dim usedLength As Long
    '
    'Compute usable length
    If startIndex + length_ > pBuffer.endIndex_ + 1 Then
        usedLength = pBuffer.endIndex_ + 1 - startIndex
    Else
        usedLength = length_
    End If
    '
    Substring = VBA.Mid$(pBuffer.text_, startIndex, usedLength)
End Function


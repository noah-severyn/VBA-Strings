Attribute VB_Name = "Strings"
'@Folder("VBA-Strings")
'@IgnoreModule ProcedureNotUsed
'Version(0.1)
Option Explicit

'compare
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


Public Function Contains(ByVal searchString As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    Contains = IndexOf(searchString, stringToFind, , compare) >= 0
End Function



Public Function EndsWith(ByVal stringToSearch As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    If Comparison = 0 Then
        EndsWith = Right$(stringToSearch, Len(stringToFind)) = stringToFind
    Else
        EndsWith = IndexOfBetween(stringToSearch, stringToFind, Len(stringToSearch) - Len(stringToFind), Len(stringToFind), compare) >= 0
    End If
End Function



Public Function Equals(ByVal baseString As String, ByVal compareString As String) As Boolean
    Equals = baseString = compareString
End Function



'0-based index
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


Public Function IndexOfBetween(ByVal searchString As String, ByVal stringToFind As String, ByVal startPos As Long, ByVal length As Long, Optional ByVal compare As Comparison = 0) As Long
    If compare Mod 2 = 1 Then
        searchString = LCase$(searchString)
        stringToFind = LCase$(stringToFind)
    End If
    
    If compare = Binary Or BinaryIgnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, length), stringToFind, vbBinaryCompare) - 1
    ElseIf compare = Text Or TextIngnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, length), stringToFind, vbTextCompare) - 1
    ElseIf compare = Database Or DatabaseIgnoreCase Then
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, length), stringToFind, vbDatabaseCompare) - 1
    Else
        IndexOfBetween = InStr(1, Mid$(searchString, startPos, length), stringToFind) - 1
    End If
End Function




Public Function IndexOfAny(ByVal searchString As String, ByRef stringsToFind() As String, Optional ByVal startPos As Long = 0, Optional ByVal length As Long) As Long
    
End Function

Public Function IndexOfAnyFromCollection(ByVal searchString As String, ByVal stringsToFind As Collection, Optional ByVal startPos As Long = 0, Optional ByVal length As Long) As Long

End Function




Public Function StartsWith(ByVal stringToSearch As String, ByVal stringToFind As String, Optional ByVal compare As Comparison = 0) As Boolean
    If Comparison = 0 Then
        StartsWith = Left$(stringToSearch, Len(stringToFind)) = stringToFind
    Else
        StartsWith = IndexOfBetween(stringToSearch, stringToFind, 0, Len(stringToFind), compare) >= 0
    End If
End Function




Public Function Substring(ByVal stringToCut As String, ByVal atPosition As Long, Optional ByVal length As Long = -1) As String
    If length = -1 Then
        Substring = Mid$(stringToCut, atPosition)
    Else
        Substring = Mid$(stringToCut, atPosition, length)
    End If
    
End Function




Public Function ToLower(ByVal stringToLowercase As String) As String
    ToLower = LCase$(stringToLowercase)
End Function




Public Function ToUpper(ByVal stringToUppercase As String) As String
    ToUpper = UCase$(stringToUppercase)
End Function



# .NET String Functions
The following table outlines all of the constructors, properties, fields, and functions in the .NET [String](https://learn.microsoft.com/en-us/dotnet/api/system.string?view=net-7.0) class and their equivalents in this repository. .NET functions with an ❌ have no equivalent in this repository.

## Constructors
| Constructor | Description | `Strings.bas` Name |
| :---------- | :---------- | :----------------------- |
| String(Char*)  | Initializes a new instance of the String class to the value indicated by a specified pointer to an array of Unicode characters. | ❌ |
| String(Char*, Int32, Int32)  | Initializes a new instance of the String class to the value indicated by a specified pointer to an array of Unicode characters, a starting character position within that array, and a length. | ❌ |
| String(Char, Int32) | Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times. | ❌ |
| String(Char[]) | Initializes a new instance of the String class to the Unicode characters indicated in the specified character array. | ❌ |
| String(Char[], Int32, Int32) | Initializes a new instance of the String class to the value indicated by an array of Unicode characters, a starting character position within that array, and a length. | ❌ |
| String(ReadOnlySpan\<Char>) | Initializes a new instance of the String class to the Unicode characters indicated in the specified read-only span. | ❌ |
| String(SByte*) | Initializes a new instance of the String class to the value indicated by a pointer to an array of 8-bit signed integers. | ❌ |
| String(SByte*, Int32, Int32) | Initializes a new instance of the String class to the value indicated by a specified pointer to an array of 8-bit signed integers, a starting position within that array, and a length. | ❌ |
| String(SByte*, Int32, Int32, Encoding) | Initializes a new instance of the String class to the value indicated by a specified pointer to an array of 8-bit signed integers, a starting position within that array, a length, and an Encoding object. | ❌ |

## Properties
| Property Name | Description | `Strings.bas` Name |
| :------------ | :---------- | :----------------------- |
| Chars[Int32]  | Gets the Char object at a specified position in the current String object. | Char |
| Length  | GGets the number of characters in the current String object. | Length |

## Fields
| Field Name | Description | `Strings.bas` Name |
| :--------- | :---------- | :----------------------- |
| Empty  | Represents the empty string. This field is read-only. | EmptyString |

## Functions
| Function Name | Description | `Strings.bas` Name |
| :------------ | :---------- | :----------------------- |
| Clone()  | Returns a reference to this instance of String. | ❌ | ❌|
| Compare(String, Int32, String, Int32, Int32)  | Compares substrings of two specified String objects and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| Compare(String, Int32, String, Int32, Int32, Boolean)  | Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| Compare(String, Int32, String, Int32, Int32, Boolean, CultureInfo)  | Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| Compare(String, Int32, String, Int32, Int32, CultureInfo, CompareOptions)  | Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order. | ❌ | ❌|
| Compare(String, Int32, String, Int32, Int32, StringComparison)  | Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| Compare(String, String)  | Compares two specified String objects and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| Compare(String, String, Boolean)  | Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| Compare(String, String, Boolean, CultureInfo)  | Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| Compare(String, String, CultureInfo, CompareOptions)  | Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order. | ❌ | ❌|
| Compare(String, String, StringComparison)  | Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order. | ❌ | ❌|
| CompareOrdinal(String, Int32, String, Int32, Int32)  | Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring. | ❌ | ❌|
| CompareOrdinal(String, String)  | Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string. | ❌ | ❌|
| CompareTo(Object)  | Compares this instance with a specified Object and indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified Object. | ❌ | ❌|
| CompareTo(String)  | Compares this instance with a specified String object and indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified string. | ❌ | ❌|
| Concat(IEnumerable\<String>)  | Concatenates the members of a constructed IEnumerable<T> collection of type String. | Concat | Append|
| Concat(Object)  | Creates the string representation of a specified object. | Concat | Append|
| Concat(Object, Object)  | Concatenates the string representations of two specified objects. | Concat | Append|
| Concat(Object, Object, Object)  | Concatenates the string representations of three specified objects. | Concat | Append|
| Concat(Object[])  | Concatenates the string representations of the elements in a specified Object array. | Concat | AppendMultiple|
| Concat(ReadOnlySpan\<Char>, ReadOnlySpan\<Char>)  | Concatenates the string representations of two specified read-only character spans. | Concat | AppendMultiple|
| Concat(ReadOnlySpan\<Char>, ReadOnlySpan\<Char>, ReadOnlySpan\<Char>)  | Concatenates the string representations of three specified read-only character spans. | Concat | AppendMultiple|
| Concat(ReadOnlySpan\<Char>, ReadOnlySpan\<Char>, ReadOnlySpan\<Char>, ReadOnlySpan\<Char>)  | Concatenates the string representations of four specified read-only character spans. | Concat | AppendMultiple|
| Concat(String, String)  | Concatenates two specified instances of String. | Concat | AppendMultiple|
| Concat(String, String, String)  | Concatenates three specified instances of String. | Concat | AppendMultiple|
| Concat(String, String, String, String)  | Concatenates four specified instances of String. | Concat | AppendMultiple|
| Concat(String[])  | Concatenates the elements of a specified String array. | Concat | AppendMultiple|
| Concat<T>(IEnumerable<T>)  | Concatenates the members of an IEnumerable<T> implementation. | Concat | AppendMultiple|
| Contains(Char)  | Returns a value indicating whether a specified character occurs within this string. | Contains | |
| Contains(Char, StringComparison)  | Returns a value indicating whether a specified character occurs within this string, using the specified comparison rules. | Contains | |
| Contains(String)  | Returns a value indicating whether a specified substring occurs within this string. | Contains | |
| Contains(String, StringComparison)  | Returns a value indicating whether a specified string occurs within this string, using the specified comparison rules. | Contains | |
| CopyTo(Int32, Char[], Int32, Int32)  | Copies a specified number of characters from a specified position in this instance to a specified position in an array of Unicode characters. | CopyToCharArrayFrom | |
| CopyTo(Span\<Char>)  | Copies the contents of this string into the destination span. | CopyToCharArray | |
| Create(IFormatProvider, DefaultInterpolatedStringHandler)  | Creates a new string by using the specified provider to control the formatting of the specified interpolated string. | ❌ | ❌|
| Create(IFormatProvider, Span\<Char>, DefaultInterpolatedStringHandler)  | Creates a new string by using the specified provider to control the formatting of the specified interpolated string. | ❌ | ❌|
| Create\<TState>(Int32, TState, SpanAction\<Char,TState>)  | Creates a new string with a specific length and initializes it after creation by using the specified callback. | Create | |
| EndsWith(Char)  | Determines whether the end of this string instance matches the specified character. | EndsWith | |
| EndsWith(String)  | Determines whether the end of this string instance matches the specified string. | EndsWith | |
| EndsWith(String, Boolean, CultureInfo)  | Determines whether the end of this string instance matches the specified string when compared using the specified culture. | EndsWith | |
| EndsWith(String, StringComparison)  | Determines whether the end of this string instance matches the specified string when compared using the specified comparison option. | EndsWith | |
| EnumerateRunes()  | Returns an enumeration of Rune from this string. | ❌ | ❌|
| Equals(Object)  | Determines whether this instance and a specified object, which must also be a String object, have the same value. | ❌ | ❌|
| Equals(String)  | Determines whether this instance and another specified String object have the same value. | Equals | |
| Equals(String, String)  | Determines whether two specified String objects have the same value. | Equals | |
| Equals(String, String, StringComparison)  | Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison. | Equals | |
| Equals(String, StringComparison)  | Determines whether this string and a specified String object have the same value. A parameter specifies the culture, case, and sort rules used in the comparison. | Equals | |
| Format(IFormatProvider, String, Object)  | Replaces the format item or items in a specified string with the string representation of the corresponding object. A parameter supplies culture-specific formatting information. |  | |
| Format(IFormatProvider, String, Object, Object)  | Replaces the format items in a string with the string representation of two specified objects. A parameter supplies culture-specific formatting information. |  | |
| Format(IFormatProvider, String, Object, Object, Object)  | Replaces the format items in a string with the string representation of three specified objects. An parameter supplies culture-specific formatting information. |  | |
| Format(IFormatProvider, String, Object[])  | Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information. |  | |
| Format(String, Object)  | Replaces one or more format items in a string with the string representation of a specified object. |  | |
| Format(String, Object, Object)  | Replaces the format items in a string with the string representation of two specified objects. |  | |
| Format(String, Object, Object, Object)  | Replaces the format items in a string with the string representation of three specified objects. |  | |
| Format(String, Object[])  | Replaces the format item in a specified string with the string representation of a corresponding object in a specified array. |  | |
| GetEnumerator()  | Retrieves an object that can iterate through the individual characters in this string. | ❌ | ❌|
| GetHashCode()  | Returns the hash code for this string. | ❌ | ❌|
| GetHashCode(ReadOnlySpan\<Char>)  | Returns the hash code for the provided read-only character span. | ❌ | ❌|
| GetHashCode(ReadOnlySpan\<Char>, StringComparison)  | Returns the hash code for the provided read-only character span using the specified rules. | ❌ | ❌|
| GetHashCode(StringComparison)  | Returns the hash code for this string using the specified rules. | ❌ | ❌|
| GetType()  | Gets the Type of the current instance. (Inherited from Object) | ❌ | ❌|
| GetTypeCode()  | Returns the TypeCode for the String class. | GetTypeCode | |
| IndexOf(Char)  | Reports the zero-based index of the first occurrence of the specified Unicode character in this string. | IndexOf | |
| IndexOf(Char, Int32)  | Reports the zero-based index of the first occurrence of the specified Unicode character in this string. The search starts at a specified character position. | IndexOf | |
| IndexOf(Char, Int32, Int32)  | Reports the zero-based index of the first occurrence of the specified character in this instance. The search starts at a specified character position and examines a specified number of character positions. | IndexOfBetween | |
| IndexOf(Char, StringComparison)  | Reports the zero-based index of the first occurrence of the specified Unicode character in this string. A parameter specifies the type of search to use for the specified character. | IndexOf | |
| IndexOf(String)  | Reports the zero-based index of the first occurrence of the specified string in this instance. | IndexOf | |
| IndexOf(String, Int32)  | Reports the zero-based index of the first occurrence of the specified string in this instance. The search starts at a specified character position. | IndexOf | |
| IndexOf(String, Int32, Int32)  | Reports the zero-based index of the first occurrence of the specified string in this instance. The search starts at a specified character position and examines a specified number of character positions. | IndexOfBetween | |
| IndexOf(String, Int32, Int32, StringComparison)  | Reports the zero-based index of the first occurrence of the specified string in the current String object. Parameters specify the starting search position in the current string, the number of characters in the current string to search, and the type of search to use for the specified string. | IndexOfBetween | |
| IndexOf(String, Int32, StringComparison)  | Reports the zero-based index of the first occurrence of the specified string in the current String object. Parameters specify the starting search position in the current string and the type of search to use for the specified string. | IndexOf | |
| IndexOf(String, StringComparison)  | Reports the zero-based index of the first occurrence of the specified string in the current String object. A parameter specifies the type of search to use for the specified string. | IndexOf | |
| IndexOfAny(Char[])  | Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters. | IndexOfAny | |
| IndexOfAny(Char[], Int32)  | Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters. The search starts at a specified character position. | IndexOfAny | |
| IndexOfAny(Char[], Int32, Int32)  | Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters. The search starts at a specified character position and examines a specified number of character positions. | IndexOfAny | |
| Insert(Int32, String)  | Returns a new string in which a specified string is inserted at a specified index position in this instance. | Insert | |
| Intern(String)  | Retrieves the system's reference to the specified String. | ❌ | ❌|
| IsInterned(String)  | Retrieves a reference to a specified String. | ❌ | ❌|
| IsNormalized()  | Indicates whether this string is in Unicode normalization form C. | ❌ | ❌|
| IsNormalized(NormalizationForm)  | Indicates whether this string is in the specified Unicode normalization form. | ❌ | ❌|
| IsNullOrEmpty(String)  | Indicates whether the specified string is null or an empty string (""). | IsNull | |
| IsNullOrWhiteSpace(String)  | Indicates whether a specified string is null, empty, or consists only of white-space characters. | IsNullOrWhiteSpace | |
| Join(Char, Object[])  | Concatenates the string representations of an array of objects, using the specified separator between each member. | Join | |
| Join(Char, String[])  | Concatenates an array of strings, using the specified separator between each member. | Join | |
| Join(Char, String[], Int32, Int32)  | Concatenates an array of strings, using the specified separator between each member, starting with the element in value located at the startIndex position, and concatenating up to count elements. | JoinBetween | |
| Join(String, IEnumerable\<String>)  | Concatenates the members of a constructed IEnumerable\<T> collection of type String, using the specified separator between each member. | Join | |
| Join(String, Object[])  | Concatenates the elements of an object array, using the specified separator between each element. | Join | |
| Join(String, String[])  | Concatenates all the elements of a string array, using the specified separator between each element. | Join | |
| Join(String, String[], Int32, Int32)  | Concatenates the specified elements of a string array, using the specified separator between each element. | JoinBetween | |
| Join\<T>(Char, IEnumerable\<T>)  | Concatenates the members of a collection, using the specified separator between each member. | Join | |
| Join\<T>(String, IEnumerable\<T>)  | Concatenates the members of a collection, using the specified separator between each member. | Join | |
| LastIndexOf(Char)  | Reports the zero-based index position of the last occurrence of a specified Unicode character within this instance. | LastIndexOf | |
| LastIndexOf(Char, Int32)  | Reports the zero-based index position of the last occurrence of a specified Unicode character within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string. | LastIndexOf | |
| LastIndexOf(Char, Int32, Int32)  | Reports the zero-based index position of the last occurrence of the specified Unicode character in a substring within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string for a specified number of character positions. | LastIndexOfBetween | |
| LastIndexOf(String)  | Reports the zero-based index position of the last occurrence of a specified string within this instance. | LastIndexOf | |
| LastIndexOf(String, Int32)  | Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string. | LastIndexOf | |
| LastIndexOf(String, Int32, Int32)  | Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string for a specified number of character positions. | LastIndexOfBetween | |
| LastIndexOf(String, Int32, Int32, StringComparison)  | Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string for the specified number of character positions. A parameter specifies the type of comparison to perform when searching for the specified string. | LastIndexOfBetween | |
| LastIndexOf(String, Int32, StringComparison)  | Reports the zero-based index of the last occurrence of a specified string within the current String object. The search starts at a specified character position and proceeds backward toward the beginning of the string. A parameter specifies the type of comparison to perform when searching for the specified string. | LastIndexOf | |
| LastIndexOf(String, StringComparison)  | Reports the zero-based index of the last occurrence of a specified string within the current String object. A parameter specifies the type of search to use for the specified string. | LastIndexOf | |
| LastIndexOfAny(Char[])  | Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode array. | LastIndexOfAny | |
| LastIndexOfAny(Char[], Int32)  | Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode array. The search starts at a specified character position and proceeds backward toward the beginning of the string. | LastIndexOfAny | |
| LastIndexOfAny(Char[], Int32, Int32)  | Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode array. The search starts at a specified character position and proceeds backward toward the beginning of the string for a specified number of character positions. | LastIndexOfAny | |
| MemberwiseClone()  | Creates a shallow copy of the current Object. (Inherited from Object) | ❌ | |
| Normalize()  | Returns a new string whose textual value is the same as this string, but whose binary representation is in Unicode normalization form C. | ❌ | |
| Normalize(NormalizationForm)  | Returns a new string whose textual value is the same as this string, but whose binary representation is in the specified Unicode normalization form. | ❌ | |
| PadLeft(Int32)  | Returns a new string that right-aligns the characters in this instance by padding them with spaces on the left, for a specified total length. | PadLeft | |
| PadLeft(Int32, Char)  | Returns a new string that right-aligns the characters in this instance by padding them on the left with a specified Unicode character, for a specified total length. | PadLeft | |
| PadRight(Int32)  | Returns a new string that left-aligns the characters in this string by padding them with spaces on the right, for a specified total length. | PadRight | |
| PadRight(Int32, Char)  | Returns a new string that left-aligns the characters in this string by padding them on the right with a specified Unicode character, for a specified total length. | PadRight | |
| Remove(Int32)  | Returns a new string in which all the characters in the current instance, beginning at a specified position and continuing through the last position, have been deleted. | Remove | |
| Remove(Int32, Int32)  | Returns a new string in which a specified number of characters in the current instance beginning at a specified position have been deleted. | Remove | |
| Replace(Char, Char)  | Returns a new string in which all occurrences of a specified Unicode character in this instance are replaced with another specified Unicode character. | Replace | |
| Replace(String, String)  | Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string. | Replace | |
| Replace(String, String, Boolean, CultureInfo)  | Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string, using the provided culture and case sensitivity. | ❌ | |
| Replace(String, String, StringComparison)  | Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string, using the provided comparison type. | Replace | |
| ReplaceLineEndings()  | Replaces all newline sequences in the current string with NewLine. | ReplaceNewLineChars | |
| ReplaceLineEndings(String)  | Replaces all newline sequences in the current string with replacementText. | ReplaceNewLineChars | |
| Split(Char, Int32, StringSplitOptions)  | Splits a string into a maximum number of substrings based on a specified delimiting character and, optionally, options. Splits a string into a maximum number of substrings based on the provided character separator, optionally omitting empty substrings from the result. |  | |
| Split(Char, StringSplitOptions)  | Splits a string into substrings based on a specified delimiting character and, optionally, options. |  | |
| Split(Char[])  | Splits a string into substrings based on specified delimiting characters. |  | |
| Split(Char[], Int32)  | Splits a string into a maximum number of substrings based on specified delimiting characters. |  | |
| Split(Char[], Int32, StringSplitOptions)  | Splits a string into a maximum number of substrings based on specified delimiting characters and, optionally, options. |  | |
| Split(Char[], StringSplitOptions)  | Splits a string into substrings based on specified delimiting characters and options. |  | |
| Split(String, Int32, StringSplitOptions)  | Splits a string into a maximum number of substrings based on a specified delimiting string and, optionally, options. |  | |
| Split(String, StringSplitOptions)  | Splits a string into substrings that are based on the provided string separator. |  | |
| Split(String[], Int32, StringSplitOptions)  | Splits a string into a maximum number of substrings based on specified delimiting strings and, optionally, options. |  | |
| Split(String[], StringSplitOptions)  | Splits a string into substrings based on a specified delimiting string and, optionally, options. |  | |
| StartsWith(Char)  | Determines whether this string instance starts with the specified character. | StartsWith | |
| StartsWith(String)  | Determines whether the beginning of this string instance matches the specified string. | StartsWith | |
| StartsWith(String, Boolean, CultureInfo)  | Determines whether the beginning of this string instance matches the specified string when compared using the specified culture. | StartsWith | |
| StartsWith(String, StringComparison)  | Determines whether the beginning of this string instance matches the specified string when compared using the specified comparison option. | StartsWith | |
| Substring(Int32)  | Retrieves a substring from this instance. The substring starts at a specified character position and continues to the end of the string. | Substring | |
| Substring(Int32, Int32)  | Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length. | Substring | |
| ToCharArray()  | Copies the characters in this instance to a Unicode character array. | ToCharArray | |
| ToCharArray(Int32, Int32)  | Copies the characters in a specified substring in this instance to a Unicode character array. | ToCharArray | |
| ToLower()  | Returns a copy of this string converted to lowercase. | ToLower | |
| ToLower(CultureInfo)  | Returns a copy of this string converted to lowercase, using the casing rules of the specified culture. | ToLower | |
| ToLowerInvariant()  | Returns a copy of this String object converted to lowercase using the casing rules of the invariant culture. | ToLower | |
| ToString()  | Returns this instance of String; no actual conversion is performed. | ❌ | ToString|
| ToString(IFormatProvider)  | Returns this instance of String; no actual conversion is performed. | ❌ | ToString|
| ToUpper()  | Returns a copy of this string converted to uppercase. | ToUpper | |
| ToUpper(CultureInfo)  | Returns a copy of this string converted to uppercase, using the casing rules of the specified culture. | ToUpper | |
| ToUpperInvariant()  | Returns a copy of this String object converted to uppercase using the casing rules of the invariant culture. | ToUpper | |
| Trim()  | Removes all leading and trailing white-space characters from the current string. | Trim | |
| Trim(Char)  | Removes all leading and trailing instances of a character from the current string. | Trim | |
| Trim(Char[])  | Removes all leading and trailing occurrences of a set of characters specified in an array from the current string. | Trim | |
| TrimEnd()  | Removes all the trailing white-space characters from the current string. | TrimEnd | |
| TrimEnd(Char)  | Removes all the trailing occurrences of a character from the current string. | TrimEnd | |
| TrimEnd(Char[])  | Removes all the trailing occurrences of a set of characters specified in an array from the current string. | TrimEnd | |
| TrimStart()  | Removes all the leading white-space characters from the current string. | TrimStart | |
| TrimStart(Char)  | Removes all the leading occurrences of a specified character from the current string. | TrimStart | |
| TrimStart(Char[])  | Removes all the leading occurrences of a set of characters specified in an array from the current string. | TrimStart | |
| TryCopyTo(Span\<Char>)  | Copies the contents of this string into the destination span. | ❌ | ❌|

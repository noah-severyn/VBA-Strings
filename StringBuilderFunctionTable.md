# .NET StringBuilder Functions
The following table outlines all of the constructors, properties, fields, and functions in the .NET [StringBuilder](https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=net-7.0) class and their equivalents in this repository. .NET functions with an ❌ have no equivalent in this repository.

| Constructor | Description | `Stringbuilder.cls` Name |
| :---------- | :---------- | :----------------------- |
| StringBuilder()  | Initializes a new instance of the StringBuilder class. | ❌ |
| StringBuilder(Int32)  | Initializes a new instance of the StringBuilder class using the specified capacity. | ❌ |
| StringBuilder(Int32, Int32)  | Initializes a new instance of the StringBuilder class that starts with a specified capacity and can grow to a specified maximum. | ❌ |
| StringBuilder(String)  | Initializes a new instance of the StringBuilder class using the specified string. | ❌ |
| StringBuilder(String, Int32)  | Initializes a new instance of the StringBuilder class using the specified string and capacity. | ❌ |
| StringBuilder(String, Int32, Int32, Int32)  | Initializes a new instance of the StringBuilder class from the specified substring and capacity. | ❌ |

| Property Name | Description | `Stringbuilder.cls` Name |
| :------------ | :---------- | :----------------------- |
| Capacity  | Gets or sets the maximum number of characters that can be contained in the memory allocated by the current instance. | Capacity |
| Chars[Int32]  | Gets or sets the character at the specified character position in this instance. | Char |
| Length  | Gets or sets the length of the current StringBuilder object. | Length |
| MaxCapacity  | Gets the maximum capacity of this instance. | MaxCapacity |

| Property Name | Description | `Stringbuilder.cls` Name |
| :------------ | :---------- | :----------------------- |
| Append(Boolean)  | Appends the string representation of a specified Boolean value to this instance. | Append |
| Append(Byte)  | Appends the string representation of a specified 8-bit unsigned integer to this instance. | Append |
| Append(Char)  | Appends the string representation of a specified Char object to this instance. | Append |
| Append(Char*, Int32)  | Appends an array of Unicode characters starting at a specified address to this instance. | ❌ |
| Append(Char, Int32)  | Appends a specified number of copies of the string representation of a Unicode character to this instance. | ❌ |
| Append(Char[])  | Appends the string representation of the Unicode characters in a specified array to this instance. | AppendMultiple |
| Append(Char[], Int32, Int32)  | Appends the string representation of a specified subarray of Unicode characters to this instance. | ❌ |
| Append(Decimal)  | Appends the string representation of a specified decimal number to this instance. | Append |
| Append(Double)  | Appends the string representation of a specified double-precision floating-point number to this instance. | Append |
| Append(IFormatProvider, StringBuilder+AppendInterpolatedStringHandler)  | Appends the specified interpolated string to this instance using the specified format. | ❌ |
| Append(Int16)  | Appends the string representation of a specified 16-bit signed integer to this instance. | Append |
| Append(Int32)  | Appends the string representation of a specified 32-bit signed integer to this instance. | Append |
| Append(Int64)  | Appends the string representation of a specified 64-bit signed integer to this instance. | Append |
| Append(Object)  | Appends the string representation of a specified object to this instance. | ❌ |
| Append(ReadOnlyMemory\<Char>)  | Appends the string representation of a specified read-only character memory region to this instance. | ❌ |
| Append(ReadOnlySpan\<Char>)  | Appends the string representation of a specified read-only character span to this instance. | ❌ |
| Append(SByte)  | Appends the string representation of a specified 8-bit signed integer to this instance. | Append |
| Append(Single)  | Appends the string representation of a specified single-precision floating-point number to this instance. | Append |
| Append(String)  | Appends a copy of the specified string to this instance. | Append |
| Append(String, Int32, Int32)  | Appends a copy of a specified substring to this instance. | ❌ |
| Append(StringBuilder)  | Appends the string representation of a specified string builder to this instance. | Append |
| Append(StringBuilder, Int32, Int32)  | Appends a copy of a substring within a specified string builder to this instance. | ❌ |
| Append(StringBuilder+AppendInterpolatedStringHandler)  | Appends the specified interpolated string to this instance. | ❌ |
| Append(UInt16)  | Appends the string representation of a specified 16-bit unsigned integer to this instance. | Append |
| Append(UInt32)  | Appends the string representation of a specified 32-bit unsigned integer to this instance. | Append |
| Append(UInt64)  | Appends the string representation of a specified 64-bit unsigned integer to this instance. | Append |
| AppendFormat(IFormatProvider, String, Object)  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a single argument using a specified format provider. | ❌ |
| AppendFormat(IFormatProvider, String, Object, Object)  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of two arguments using a specified format provider. | ❌ |
| AppendFormat(IFormatProvider, String, Object, Object, Object)  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of three arguments using a specified format provider. | ❌ |
| AppendFormat(IFormatProvider, String, Object[])  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a corresponding argument in a parameter array using a specified format provider. | ❌ |
| AppendFormat(String, Object)  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a single argument. | ❌ |
| AppendFormat(String, Object, Object)  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of two arguments. | ❌ |
| AppendFormat(String, Object, Object, Object)  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of three arguments. | ❌ |
| AppendFormat(String, Object[])  | Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a corresponding argument in a parameter array. | ❌ |
| AppendJoin(Char, Object[])  | Concatenates the string representations of the elements in the provided array of objects, using the specified char separator between each member, then appends the result to the current instance of the string builder. | AppendJoin |
| AppendJoin(Char, String[])  | Concatenates the strings of the provided array, using the specified char separator between each string, then appends the result to the current instance of the string builder. | AppendJoin |
| AppendJoin(String, Object[])  | Concatenates the string representations of the elements in the provided array of objects, using the specified separator between each member, then appends the result to the current instance of the string builder. | AppendJoin |
| AppendJoin(String, String[])  | Concatenates the strings of the provided array, using the specified separator between each string, then appends the result to the current instance of the string builder. | AppendJoin |
| AppendJoin\<T>(Char, IEnumerable\<T>)  | Concatenates and appends the members of a collection, using the specified char separator between each member. | AppendJoin |
| AppendJoin\<T>(String, IEnumerable\<T>)  | Concatenates and appends the members of a collection, using the specified separator between each member. | AppendJoin |
| AppendLine()  | Appends the default line terminator to the end of the current StringBuilder object. | AppendLine |
| AppendLine(IFormatProvider, StringBuilder+AppendInterpolatedStringHandler)  | Appends the specified interpolated string using the specified format, followed by the default line terminator, to the end of the current StringBuilder object. | AppendLine |
| AppendLine(String)  | Appends a copy of the specified string followed by the default line terminator to the end of the current StringBuilder object. | AppendLine |
| AppendLine(StringBuilder+AppendInterpolatedStringHandler)  | Appends the specified interpolated string followed by the default line terminator to the end of the current StringBuilder object. | AppendLine |
| Clear()  | Removes all characters from the current StringBuilder instance. | Clear |
| CopyTo(Int32, Char[], Int32, Int32)  | Copies the characters from a specified segment of this instance to a specified segment of a destination Char array. | ❌ |
| CopyTo(Int32, Span\<Char>, Int32)  | Copies the characters from a specified segment of this instance to a destination Char span. | ❌ |
| EnsureCapacity(Int32)  | Ensures that the capacity of this instance of StringBuilder is at least the specified value. | EnsureCapacity |
| Equals(Object)  | Determines whether the specified object is equal to the current object. (Inherited from Object) | ❌ |
| Equals(ReadOnlySpan\<Char>)  | Returns a value indicating whether the characters in this instance are equal to the characters in a specified read-only character span. | ❌ |
| Equals(StringBuilder)  | Returns a value indicating whether this instance is equal to a specified object. | ❌ |
| GetChunks()  | Returns an object that can be used to iterate through the chunks of characters represented in a ReadOnlyMemory\<Char> created from this StringBuilder instance. | ❌ |
| GetHashCode()  | Serves as the default hash function. (Inherited from Object) | ❌ |
| Insert(Int32, Boolean)  | Inserts the string representation of a Boolean value into this instance at the specified character position. | Insert |
| Insert(Int32, Byte)  | Inserts the string representation of a specified 8-bit unsigned integer into this instance at the specified character position. | Insert |
| Insert(Int32, Char)  | Inserts the string representation of a specified Unicode character into this instance at the specified character position. | Insert |
| Insert(Int32, Char[])  | Inserts the string representation of a specified array of Unicode characters into this instance at the specified character position. | ❌ |
| Insert(Int32, Char[], Int32, Int32)  | Inserts the string representation of a specified subarray of Unicode characters into this instance at the specified character position. | ❌ |
| Insert(Int32, Decimal)  | Inserts the string representation of a decimal number into this instance at the specified character position. | Insert |
| Insert(Int32, Double)  | Inserts the string representation of a double-precision floating-point number into this instance at the specified character position. | Insert |
| Insert(Int32, Int16)  | Inserts the string representation of a specified 16-bit signed integer into this instance at the specified character position. | Insert |
| Insert(Int32, Int32)  | Inserts the string representation of a specified 32-bit signed integer into this instance at the specified character position. | Insert |
| Insert(Int32, Int64)  | Inserts the string representation of a 64-bit signed integer into this instance at the specified character position. | Insert |
| Insert(Int32, Object)  | Inserts the string representation of an object into this instance at the specified character position. | ❌ |
| Insert(Int32, ReadOnlySpan\<Char>)  | Inserts the sequence of characters into this instance at the specified character position. | ❌ |
| Insert(Int32, SByte)  | Inserts the string representation of a specified 8-bit signed integer into this instance at the specified character position. | Insert |
| Insert(Int32, Single)  | Inserts the string representation of a single-precision floating point number into this instance at the specified character position. | Insert |
| Insert(Int32, String)  | Inserts a string into this instance at the specified character position. | Insert |
| Insert(Int32, String, Int32)  | Inserts one or more copies of a specified string into this instance at the specified character position. | Insert |
| Insert(Int32, UInt16)  | Inserts the string representation of a 16-bit unsigned integer into this instance at the specified character position. | Insert |
| Insert(Int32, UInt32)  | Inserts the string representation of a 32-bit unsigned integer into this instance at the specified character position. | Insert |
| Insert(Int32, UInt64)  | Inserts the string representation of a 64-bit unsigned integer into this instance at the specified character position. | Insert |
| MemberwiseClone()  | Creates a shallow copy of the current Object. (Inherited from Object) | ❌ |
| Remove(Int32, Int32)  | Removes the specified range of characters from this instance. | Remove |
| Replace(Char, Char)  | Replaces all occurrences of a specified character in this instance with another specified character. | Replace |
| Replace(Char, Char, Int32, Int32)  | Replaces, within a substring of this instance, all occurrences of a specified character with another specified character. | Replace |
| Replace(String, String)  | Replaces all occurrences of a specified string in this instance with another specified string. | Replace |
| Replace(String, String, Int32, Int32)  | Replaces, within a substring of this instance, all occurrences of a specified string with another specified string. | Overwrite |
| ToString()  | Converts the value of this instance to a String. | ToString |
| ToString(Int32, Int32)  | Converts the value of a substring of this instance to a String. | ToString |

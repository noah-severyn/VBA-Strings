# String Builder
## About
**StringBuilder** is a class that provides a string-like object which offers superior performance when extensive
string manipulation operations are required (such as modifying a string numerous times in a loop). Modifying a string repeatedly can exact a significant performance penalty - see more information [here](https://ramblings.mcpher.com/optimization-links/strings-and-garbage/) and [here](https://ramblings.mcpher.com/strings-and-garbage-collector-in-vba/). 

The alternative is to use StringBuilder, which is a mutable string class. Mutability means that once an instance of the class has been created, it can be modified by appending, removing, replacing, or inserting characters. Strings are immutable, meaning that a modification to the string actually returns a *new* string with the modifications - the base string cannot be modified.

A StringBuilder object maintains a buffer to accommodate expansions to the string. New data is appended to the buffer if room is available; otherwise, a new, larger buffer is allocated, data from the original buffer is copied to the new buffer, and the new data is then appended to the new buffer. This means that operations can be applied to the string *in place*, instead of creating a new string each operation.

The implementation is strongly based on Christian Buse's [VBA-StringBuffer](https://github.com/cristianbuse/VBA-StringBuffer), which is the best performing VBA StringBuilder I have found based on [limited testing](/docs/TestingResults.md).

## Usage Explanation
*The following paragraphs are adapted from the [.NET StringBuilder documentation](https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=net-7.0#StringAndSB).*

> **Note**
> Although the StringBuilder class generally offers better performance than the String class, you should not automatically replace String with StringBuilder whenever you want to manipulate strings. Performance depends on the size of the string, the amount of memory to be allocated for the new string, the system on which your code is executing, and the type of operation. You should be prepared to test your code to determine whether StringBuilder actually offers a significant performance improvement.

Consider using the **String** class under these conditions:
* When the number of changes that your code will make to a string is small (typically less than 1000). In these cases, StringBuilder might offer negligible or no performance improvement over String.
* When you are performing a fixed number of concatenation operations, particularly with string literals. In this case, the compiler might combine the concatenation operations into a single operation.
* When you have to perform extensive search operations while you are building your string. The StringBuilder class lacks search methods such as IndexOf or StartsWith. You'll have to convert the StringBuilder object to a String for these operations, and this can negate the performance benefit from using StringBuilder. For more information, see the Searching the text in a StringBuilder object section.

Consider using the **StringBuilder** class under these conditions:
* When you expect your code to make an unknown number of changes to a string at design time (for example, when you are using a loop to concatenate a random number of strings that contain user input).
* When you expect your code to make a significant number of changes to a string (typically more than 1000). 

For a full list of functions and how they line up with the built-in .NET functions, refer to the [StringBuilder function table](/docs/StringBuilderFunctionTable.md).

## Usage Examples
To use this class, import the `StringBuilder.cls` module into your project. To create a **StringBuilder** object,
```VB
Dim sb As StringBuilder
Set sb = new StringBuilder
```
Use the methods available to the class by referencing that instance
```VB
sb.Append "new text" '-> "new text"
sb.Insert 4, "inserted " '-> "new inserted text"
sb.Remove 0, 4 '->"inserted text"
```
To return the final string, use
```VB
sb.ToString '-> Returns "inserted text"
```

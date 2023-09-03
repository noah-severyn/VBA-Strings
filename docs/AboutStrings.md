# Strings
## About
**Strings** is an collection of "static" classes designed for easy and flexible string manipulation. For a full list of functions and how they line up with the built-in .NET functions, refer to the [Strings function table](/docs/StringsFunctionTable.md).

## Usage Examples
To use this class, import the `Strings.bas` module into your project. Each function requires you pass a string to manipulate on.
```VB
Strings.IndexOf("test string", "s") '-> Returns 2

stringstoFind = Array("one","two","three","four")
Strings.StartsWithAny("fourth test string", stringsToFind) '-> Returns TRUE
```
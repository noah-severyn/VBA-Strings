# VBA-Strings
## About
**VBA-String** is the most comprehensive library to-date for string manipulation in VBA and includes two modules: 
* `Strings.bas` is an collection of string manipulation functions based on the ones provided by the .NET built-in [String](https://learn.microsoft.com/en-us/dotnet/api/system.string?view=net-7.0) class.
* `StringBuilder.cls` is a custom class which offers superior performance when extensive string manipulation operations are required, and is based on the .NET built-in [StringBuilder](https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder) class.

All functions are additionally designed to use 0-based strings, providing an improved syntax in line with other languages.

## Usage and Documentation
Refer to the [About Strings](/docs/AboutStrings.md) and the [About StringBuilder](/docs/AboutStringBuilder.md) documentation.


## Inspiration & Reference
The functions in this library are based on .NET's **[String](https://learn.microsoft.com/en-us/dotnet/api/system.string?view=net-7.0)** and **[StringBuilder](https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=net-7.0)** classes. This library also incorporates  a multitude of string functions included in other VBA string-related repositories, including Christian Buse's [VBA-StringBuffer](https://github.com/cristianbuse/VBA-StringBuffer), Robert Todar's [VBA-Strings](https://github.com/todar/VBA-Strings), Bruce Mcpherson's [cStringChunker](https://gist.github.com/brucemcpherson/5102369), Peter Roach's [clsStringBuilder](https://github.com/PeterRoach/VBA/tree/main/clsStringBuilder) and [modString](https://github.com/PeterRoach/VBA/tree/main/modString), Daniele Giaquinto's [StringType](https://github.com/exSnake/VBTools/blob/master/StringType.cls), Frank Schwab's [StringBuilder](https://github.com/xformerfhs/VBAUtilities/tree/master/StringHandling), and Greedquest's [VBA-Gems](https://github.com/Greedquest/VBA-Gems).

For a full list of functions and how they line up with the functions from these other libraries, refer to the:
* [Strings function table](/docs/StringsFunctionTable.md)
* [StringBuilder function table](/docs/StringBuilderFunctionTable.md)
* [Other Libraries compatibility table](/docs/OtherReposFunctionTable.md)
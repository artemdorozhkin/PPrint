# PPrint for VBA – Pretty-Print Debugging Utility

**PPrint** is a lightweight, extensible pretty-printer for **VBA (Visual Basic for Applications)** that provides enhanced debugging output for arrays, dictionaries, ranges, collections, strings, dates, and custom objects.

## Features

- Print any number of values with a single call.
- Supports native VBA types like **String**, **Date**, **Array**, **Range**, **Collection**, and **Scripting.Dictionary**.
- Automatically formats **user-defined objects** if it implements a `Repr__` method.
- Intelligent output: clean, readable, and designed to save time during development.

## Installation

### With [ppm](https://github.com/artemdorozhkin/ppm.git)

Run from the Immediate Window:

```vba
ppm "install pprint"
```

### Manually

1. Open **Excel VBA Editor** (`ALT + F11`)
2. Go to **File > Export File...** (`Ctrl + E`)
3. Select the `PPrintModule.bas` VBA module code
4. Save and start using the functions in your macros

## Example

```vba
Sub Example()
    Dim arr(1 To 2) As Variant
    arr(1) = "Hello"
    arr(2) = 123

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "key", arr

    PPrint "Test", Date, arr, dict
End Sub
```

Output:

```vba
'Test' #03/24/2025# ['Hello', 123] {'key': ['Hello', 123]}
```

## How It Works

`PPrint` walks through each argument using type introspection:

- Strings are wrapped in 'quotes'
- Dates use #MM/DD/YYYY# format
- Arrays and collections are recursively formatted
- Dictionaries are shown as key-value maps
- Ranges display cell addresses and values
- Custom objects use a Repr__() method if defined

## Extendable Design

To support custom classes, simply add a method `Repr__`:

```vba
Public Function Repr__() As String
    Repr__ = "<MyClass: " & Me.Name & ">"
End Function
```

PPrint will automatically detect and use it.

## Keywords (for discovery)

vba debug print, vba pretty print, vba to string, vba print dictionary, vba print array, vba debug helper, visual basic pprint, vba format output, vba custom object debug

## License

MIT License. Free to use and modify.

This repository demonstrates an uncommon VBScript bug related to the `IsEmpty` function's handling of empty strings within numeric contexts. The `IsEmpty` function, intended to check for uninitialized or empty Variant variables, behaves unexpectedly when dealing with empty strings in arithmetic operations.  This leads to incorrect results and potential runtime errors.

The `bug.vbs` file shows the issue: an empty string passed to a function expecting numbers isn't correctly identified as empty by `IsEmpty`, resulting in unexpected addition.

The solution provided in `bugSolution.vbs` illustrates how to correctly handle this scenario by explicitly checking for empty strings using the `Len` function before performing arithmetic operations.
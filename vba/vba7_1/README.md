# Visual Basic 7.1 Grammar for ANTLR4

Derived from the Visual Basic 7.1 language reference

https://msopenspecs.azureedge.net/files/MS-VBAL/%5bMS-VBAL%5d.pdf

This grammar ignores conditional-compilation statements. the vba_cc grammar can be used against vba files to analyze that portion of the code.

Line endings, whitespace, and comments are traditionally removed from parsers, but the VBA standard dictates when and how some of these are valid, so they remain. Unfortunately, this leaves open room for more inaccuracy. Please report any whitespace bugs.
## Bugs Reported to Microsoft
### Open
* class file header definition
* reserved-name keywords should be able to be used within expressions
* private and public enums missing from a parent
* Missing July from English-month-names

### Closed
* line continuation format
* reserved-type-identifier missing from reserved-identifier
* VB_invoke_propertyPutRef combined with VB_MemberFlags
* erase-statement missing from data-manipulation-statements
* print-statement missing from file-statement
* enum-member-list missing from enum-declaration

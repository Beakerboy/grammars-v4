# Visual Basic 7.1 Grammar for ANTLR4

Derived from the Visual Basic 7.1 language reference

https://msopenspecs.azureedge.net/files/MS-VBAL/%5bMS-VBAL%5d.pdf

This grammar ignores conditional-compilation statements. The vba_cc grammar can be used against vba files to analyze that portion of the code.

Line endings, whitespace, and comments are traditionally removed from parsers, but the VBA standard dictates when and how some of these are valid, so they remain. Unfortunately, this leaves open room for more inaccuracy. Please report any whitespace bugs.
## Bugs Reported to Microsoft
In the course of creating this parser, several bugs were found in the MS-VBAL v1.7 document, reported to Microsoft, and are listed below.
### Open - awaiting feedback from Microsoft on proper solution.
* class file header definition
* reserved-name keywords should be able to be used within expressions
* Allow multiple case-range statement in select case.
* reserved-type-identifier missing from reserved-identifier.
* The use of attributes within the code body is not addressed.
* Redim of member access espressions
* marked-file-numbers cannot be used as an input to a function.
* missing end statement

### Closed-Fixed here and planned for inclusion in next release of MS-VBAL
* line continuation format
* VB_invoke_propertyPutRef combined with VB_MemberFlags
* erase-statement missing from data-manipulation-statements
* print-statement missing from file-statement
* enum-member-list missing from enum-declaration
* private and public enums missing from a parent
* Missing July from English-month-names

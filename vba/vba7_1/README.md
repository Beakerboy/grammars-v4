# Visual Basic 7.1 Grammar for ANTLR4

Derived from the Visual Basic 7.1 language reference

https://msopenspecs.azureedge.net/files/MS-VBAL/%5bMS-VBAL%5d.pdf

This grammar ignores conditional-compilation statements. The vba_cc grammar can be used against vba files to analyze that portion of the code.

Line endings, whitespace, and comments are traditionally removed from parsers, but the VBA standard dictates when and how some of these are valid, so they remain. Unfortunately, this leaves open room for more inaccuracy. Please report any whitespace bugs.
## Bugs Reported to Microsoft
In the course of creating this parser, several bugs were found in the MS-VBAL v1.8 document, reported to Microsoft, and are listed below.
### Open - awaiting feedback from Microsoft on proper solution.
* class file header definition
* reserved-name keywords should be able to be used within expressions
* The use of attributes within the code body is not addressed.
* public types in class modules (allowed in VB6, but not VBA...however, can be imported and exported. code will fail on compiling)
* Semicolons in Debug.Print
* VB_Ext_KEY attribute.

### Closed-Fixed here and planned for inclusion in next release of MS-VBAL
* private and public enums missing from a parent
* Allow multiple case-range statements in select case.
* missing end statement
* Redim of member access expressions

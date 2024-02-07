# Visual Basic Conditional Compilation Grammar for ANTLR4

Derived from the Visual Basic 7.1 language reference

https://msopenspecs.azureedge.net/files/MS-VBAL/%5bMS-VBAL%5d.pdf

The vba_cc grammar can be used against vba files to analyze that portion of the code.

Line endings, whitespace, and comments are traditionally removed from parsers, but the VBA standard dictates when and how some of these are valid, so they remain. Unfortunately, this leaves open room for more inaccuracy. Please report any whitespace bugs.

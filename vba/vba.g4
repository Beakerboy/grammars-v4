/*
* Copyright (C) 2014 Ulrich Wolffgang <u.wol@wwu.de>
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
* GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License
* along with this program. If not, see <http://www.gnu.org/licenses/>.
*/

/*
* Visual Basic 6.0 Grammar for ANTLR4
*
* This is an approximate grammar for Visual Basic 6.0, derived
* from the Visual Basic 6.0 language reference
* http://msdn.microsoft.com/en-us/library/aa338033%28v=vs.60%29.aspx
* and tested against MSDN VB6 statement examples as well as several Visual
* Basic 6.0 code repositories.
*
* Characteristics:
*
* 1. This grammar is line-based and takes into account whitespace, so that
*    member calls (e.g. "A.B") are distinguished from contextual object calls
*    in WITH statements (e.g. "A .B").
*
* 2. Keywords can be used as identifiers depending on the context, enabling
*    e.g. "A.Type", but not "Type.B".
*
*
* Known limitations:
*
* 1. Preprocessor statements (#if, #else, ...) must not interfere with regular
*    statements.
*
* Change log:
*
* v1.4 Rubberduck
*   - renamed to VBA; goal is to support VBA, and a shorter name is more practical.
*   - added moduleDeclarations rule, moved moduleOptions there; options can now be
*     located anywhere in declarations section, without breaking the parser.
*   - added support for Option Compare Database.
*   - added support for VBA 7.0 PtrSafe attribute for Declare statements.
*   - implemented a fileNumber rule to locate identifier usages in file numbers.
*   - added support for anonymous declarations in With blocks (With New Something)
*   - blockStmt rules being sorted alphabetically was wrong. moved implicit call statement last.
*   - '!' in dictionary call statement rule gets picked up as a type hint; changed member call
*     to accept '!' as well as '.', but this complicates resolving the '!' shorthand syntax.
*   - added a subscripts rule in procedure calls, to avoid breaking the parser with
*     a function call that returns an array that is immediately accessed.
*   - added missing macroConstStmt (#CONST) rule.
*   - amended selectCaseStmt rules to support all valid syntaxes.
*   - blockStmt is now illegal in declarations section.
*   - added ON_LOCAL_ERROR token, to support legacy ON LOCAL ERROR statements.
*   - added additional typeHint? token to declareStmt, to support "Declare Function Foo$".
*   - modified WS lexer rule to correctly account for line continuations;
*   - modified multi-word lexer rules to use WS lexer token instead of ' '; this makes
*     the grammar support "Option _\n Explicit" and other keywords being specified on multiple lines.
*	- modified moduleOption rules to account for WS token in corresponding lexer rules.
*   - modified NEWLINE lexer rule to properly support instructions separator (':').
*   - tightened DATELITERAL lexer rule to the format enforced by the VBE, because "#fn: Close #"
*     in "Dim fn: fn = FreeFile: Open "filename" For Output As #fn: Close #fn" was picked up as a date literal.
*   - redefined IDENTIFIER lexer rule to support non-Latin characters (e.g. Japanese)
*   - made seekStmt, lockStmt, unlockStmt, getStmt and widthStmt accept a fileNumber (needed to support '#')
*   - fixed precompiler directives, which can now be nested. they still can't interfere with other blocks though.
*   - optional parameters can be a valueStmt.
*   - added support for Octal and Currency literals.
*   - implemented proper specs for DATELITERAL.
*   - added comments to parse tree (removes known limitation #2).
*   - macroConstStmt now allowed in blockStmt.
*   - allow type hints for parameters.
*
*======================================================================================
*
* v1.3
*	- call statement precedence
*
* v1.2
*	- refined call statements
*
* v1.1
*	- precedence of operators and of ELSE in select statements
*	- optimized member calls
*
* v1.0 Initial revision
*/

// $antlr-format alignTrailingComments true, columnLimit 150, minEmptyLines 1, maxEmptyLinesToKeep 1, reflowComments false, useTab false
// $antlr-format allowShortRulesOnASingleLine false, allowShortBlocksOnASingleLine true, alignSemicolons hanging, alignColons hanging

grammar vba;

options {
    caseInsensitive = true;
}

// module ----------------------------------

startRule
    : module EOF
    ;

module
    : WS? endOfLine* (moduleHeader endOfLine*)? moduleConfig? endOfLine* moduleAttributes? endOfLine* moduleDeclarations? endOfLine* moduleBody?
        endOfLine* WS?
    ;

moduleHeader
    : VERSION WS DOUBLELITERAL WS CLASS?
    ;

moduleConfig
    : BEGIN (WS GUID WS ambiguousIdentifier)? endOfLine* moduleConfigElement+ END
    ;

moduleConfigElement
    : ambiguousIdentifier WS? EQ WS? literal (COLON literal)? endOfLine*
    ;

moduleAttributes
    : (attributeStmt endOfLine+)+
    ;

moduleDeclarations
    : moduleDeclarationsElement (endOfLine+ moduleDeclarationsElement)* endOfLine*
    ;

moduleOption
    : OPTION_BASE WS SHORTLITERAL                  # optionBaseStmt
    | OPTION_COMPARE WS (BINARY | TEXT | DATABASE) # optionCompareStmt
    | OPTION_EXPLICIT                              # optionExplicitStmt
    | OPTION_PRIVATE_MODULE                        # optionPrivateModuleStmt
    ;

moduleDeclarationsElement
    : comment
    | declareStmt
    | enumerationStmt
    | eventStmt
    | constStmt
    | implementsStmt
    | variableStmt
    | moduleOption
    | typeStmt
    | deftypeStmt
    | macroStmt
    ;

macroStmt
    : macroConstStmt
    | macroIfThenElseStmt
    ;

moduleBody
    : moduleBodyElement (endOfLine+ moduleBodyElement)* endOfLine*
    ;

moduleBodyElement
    : functionStmt
    | propertyGetStmt
    | propertySetStmt
    | propertyLetStmt
    | subStmt
    | macroStmt
    ;

// block ----------------------------------

attributeStmt
    : ATTRIBUTE WS implicitCallStmt_InStmt WS? EQ WS? literal (WS? ',' WS? literal)*
    ;

block
    : blockStmt (endOfStatement blockStmt)* endOfStatement
    ;

blockStmt
    : lineLabel
    | appactivateStmt
    | attributeStmt
    | beepStmt
    | chdirStmt
    | chdriveStmt
    | closeStmt
    | constStmt
    | dateStmt
    | deleteSettingStmt
    | doLoopStmt
    | endStmt
    | eraseStmt
    | errorStmt
    | exitStmt
    | explicitCallStmt
    | filecopyStmt
    | forEachStmt
    | forNextStmt
    | getStmt
    | goSubStmt
    | goToStmt
    | ifThenElseStmt
    | implementsStmt
    | inputStmt
    | killStmt
    | letStmt
    | lineInputStmt
    | lineNumber
    | loadStmt
    | lockStmt
    | lsetStmt
    | macroStmt
    | midStmt
    | mkdirStmt
    | nameStmt
    | onErrorStmt
    | onGoToStmt
    | onGoSubStmt
    | openStmt
    | printStmt
    | putStmt
    | raiseEventStmt
    | randomizeStmt
    | redimStmt
    | resetStmt
    | resumeStmt
    | returnStmt
    | rmdirStmt
    | rsetStmt
    | savepictureStmt
    | saveSettingStmt
    | seekStmt
    | selectCaseStmt
    | sendkeysStmt
    | setattrStmt
    | setStmt
    | stopStmt
    | timeStmt
    | unloadStmt
    | unlockStmt
    | variableStmt
    | whileWendStmt
    | widthStmt
    | withStmt
    | writeStmt
    | implicitCallStmt_InBlock
    | implicitCallStmt_InStmt
    ;

// statements ----------------------------------

appactivateStmt
    : APPACTIVATE WS expression (WS? ',' WS? expression)?
    ;

beepStmt
    : BEEP
    ;

chdirStmt
    : CHDIR WS expression
    ;

chdriveStmt
    : CHDRIVE WS expression
    ;

closeStmt
    : CLOSE (WS fileNumber (WS? ',' WS? fileNumber)*)?
    ;

constStmt
    : (visibility WS)? CONST WS constSubStmt (WS? ',' WS? constSubStmt)*
    ;

constSubStmt
    : ambiguousIdentifier typeHint? (WS asTypeClause)? WS? EQ WS? expression
    ;

dateStmt
    : DATE WS? EQ WS? expression
    ;

declareStmt
    : (visibility WS)? DECLARE WS (PTRSAFE WS)? ((FUNCTION typeHint?) | SUB) WS ambiguousIdentifier typeHint? WS LIB WS STRINGLITERAL (
        WS ALIAS WS STRINGLITERAL
    )? (WS? argList)? (WS asTypeClause)?
    ;

deftypeStmt
    : (
        DEFBOOL
        | DEFBYTE
        | DEFINT
        | DEFLNG
        | DEFCUR
        | DEFSNG
        | DEFDBL
        | DEFDEC
        | DEFDATE
        | DEFSTR
        | DEFOBJ
        | DEFVAR
    ) WS letterrange (WS? ',' WS? letterrange)*
    ;

deleteSettingStmt
    : DELETESETTING WS valueStmt WS? ',' WS? valueStmt (WS? ',' WS? valueStmt)?
    ;

doLoopStmt
    : DO endOfStatement block? LOOP
    | DO WS (WHILE | UNTIL) WS valueStmt endOfStatement block? LOOP
    | DO endOfStatement block LOOP WS (WHILE | UNTIL) WS valueStmt
    ;

endStmt
    : END
    ;

enumerationStmt
    : (visibility WS)? ENUM WS ambiguousIdentifier endOfStatement enumerationStmt_Constant* END_ENUM
    ;

enumerationStmt_Constant
    : ambiguousIdentifier (WS? EQ WS? valueStmt)? endOfStatement
    ;

eraseStmt
    : ERASE WS valueStmt (',' WS? valueStmt)*?
    ;

errorStmt
    : ERROR WS valueStmt
    ;

eventStmt
    : (visibility WS)? EVENT WS ambiguousIdentifier WS? argList
    ;

exitStmt
    : EXIT_DO
    | EXIT_FOR
    | EXIT_FUNCTION
    | EXIT_PROPERTY
    | EXIT_SUB
    ;

filecopyStmt
    : FILECOPY WS valueStmt WS? ',' WS? valueStmt
    ;

forEachStmt
    : FOR WS EACH WS ambiguousIdentifier typeHint? WS IN WS valueStmt endOfStatement block? NEXT (
        WS ambiguousIdentifier
    )?
    ;

forNextStmt
    : FOR WS ambiguousIdentifier typeHint? (WS asTypeClause)? WS? EQ WS? valueStmt WS TO WS valueStmt (
        WS STEP WS valueStmt
    )? endOfStatement block? NEXT (WS ambiguousIdentifier)?
    ;

functionStmt
    : (visibility WS)? (STATIC WS)? FUNCTION WS? ambiguousIdentifier typeHint? (WS? argList)? (
        WS? asTypeClause
    )? endOfStatement block? END_FUNCTION
    ;

getStmt
    : GET WS fileNumber WS? ',' WS? ezpression? WS? ',' WS? expression
    ;

goSubStmt
    : GOSUB WS expression
    ;

goToStmt
    : GOTO WS expression
    ;

ifThenElseStmt
    : IF WS ifConditionStmt WS THEN WS blockStmt (WS ELSE WS blockStmt)? # inlineIfThenElse
    | ifBlockStmt ifElseIfBlockStmt* ifElseBlockStmt? END_IF             # blockIfThenElse
    ;

ifBlockStmt
    : IF WS ifConditionStmt WS THEN endOfStatement block?
    ;

ifConditionStmt
    : valueStmt
    ;

ifElseIfBlockStmt
    : ELSEIF WS ifConditionStmt WS THEN endOfStatement block?
    ;

ifElseBlockStmt
    : ELSE endOfStatement block?
    ;

implementsStmt
    : IMPLEMENTS WS ambiguousIdentifier
    ;

inputStmt
    : INPUT WS fileNumber (WS? ',' WS? expression)+
    ;

killStmt
    : KILL WS expression
    ;

letStmt
    : (LET WS)? implicitCallStmt_InStmt WS? (EQ | PLUS_EQ | MINUS_EQ) WS? typeHint? valueStmt typeHint?
    ;

lineInputStmt
    : LINE_INPUT WS fileNumber WS? ',' WS? expression
    ;

lineNumber
    : (INTEGERLITERAL | SHORTLITERAL) NEWLINE? ':'? NEWLINE? WS?
    ;

loadStmt
    : LOAD WS expression
    ;

lockStmt
    : LOCK WS expression (WS? ',' WS? expression (WS TO WS expression)?)?
    ;

lsetStmt
    : LSET WS implicitCallStmt_InStmt WS? EQ WS? expression
    ;

macroConstStmt
    : MACRO_CONST WS? ambiguousIdentifier WS? EQ WS? expression
    ;

macroIfThenElseStmt
    : macroIfBlockStmt macroElseIfBlockStmt* macroElseBlockStmt? MACRO_END_IF
    ;

macroIfBlockStmt
    : MACRO_IF WS? ifConditionStmt WS THEN endOfStatement (moduleDeclarations | moduleBody | block)*
    ;

macroElseIfBlockStmt
    : MACRO_ELSEIF WS? ifConditionStmt WS THEN endOfStatement (
        moduleDeclarations
        | moduleBody
        | block
    )*
    ;

macroElseBlockStmt
    : MACRO_ELSE endOfStatement (moduleDeclarations | moduleBody | block)*
    ;

midStmt
    : MID WS? LPAREN WS? argsCall WS? RPAREN
    ;

mkdirStmt
    : MKDIR WS expression
    ;

nameStmt
    : NAME WS expressiom WS AS WS exprrssion
    ;

// 5.6.8
newExpression
    : NEW WS? expression
    ;

onErrorStmt
    : (ON_ERROR | ON_LOCAL_ERROR) WS (GOTO WS expression | RESUME WS NEXT)
    ;

onGoToStmt
    : ON WS expression WS GOTO WS expression (WS? ',' WS? expression)*
    ;

onGoSubStmt
    : ON WS expression WS GOSUB WS expression (WS? ',' WS? expression)*
    ;

openStmt
    : OPEN WS expression WS FOR WS (APPEND | BINARY | INPUT | OUTPUT | RANDOM) (
        WS ACCESS WS (READ | WRITE | READ_WRITE)
    )? (WS (SHARED | LOCK_READ | LOCK_WRITE | LOCK_READ_WRITE))? WS AS WS fileNumber (
        WS LEN WS? EQ WS? expression
    )?
    ;

outputList
    : outputList_Expression (WS? (';' | ',') WS? outputList_Expression?)*
    | outputList_Expression? (WS? (';' | ',') WS? outputList_Expression?)+
    ;

outputList_Expression
    : expression
    | (SPC | TAB) (WS? LPAREN WS? argsCall WS? RPAREN)?
    ;

printStmt
    : PRINT WS fileNumber WS? ',' (WS? outputList)?
    ;

propertyGetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_GET WS ambiguousIdentifier typeHint? (WS? argList)? (
        WS asTypeClause
    )? endOfStatement block? END_PROPERTY
    ;

propertySetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_SET WS ambiguousIdentifier (WS? argList)? endOfStatement block? END_PROPERTY
    ;

propertyLetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_LET WS ambiguousIdentifier (WS? argList)? endOfStatement block? END_PROPERTY
    ;

putStmt
    : PUT WS fileNumber WS? ',' WS? expression? WS? ',' WS? expression
    ;

raiseEventStmt
    : RAISEEVENT WS ambiguousIdentifier (WS? LPAREN WS? (argsCall WS?)? RPAREN)?
    ;

randomizeStmt
    : RANDOMIZE (WS expression)?
    ;

redimStmt
    : REDIM WS (PRESERVE WS)? redimSubStmt (WS? ',' WS? redimSubStmt)*
    ;

redimSubStmt
    : implicitCallStmt_InStmt WS? LPAREN WS? subscripts WS? RPAREN (WS asTypeClause)?
    ;

resetStmt
    : RESET
    ;

resumeStmt
    : RESUME (WS (NEXT | ambiguousIdentifier))?
    ;

returnStmt
    : RETURN
    ;

rmdirStmt
    : RMDIR WS expression
    ;

rsetStmt
    : RSET WS implicitCallStmt_InStmt WS? EQ WS? expression
    ;

savepictureStmt
    : SAVEPICTURE WS valueStmt WS? ',' WS? valueStmt
    ;

saveSettingStmt
    : SAVESETTING WS valueStmt WS? ',' WS? valueStmt WS? ',' WS? valueStmt WS? ',' WS? valueStmt
    ;

seekStmt
    : SEEK WS fileNumber WS? ',' WS? expression
    ;

selectCaseStmt
    : SELECT WS CASE WS expression endOfStatement sC_Case* END_SELECT
    ;

sC_Selection
    : IS WS? comparisonOperator WS? expression # caseCondIs
    | valueStmt WS TO WS expression            # caseCondTo
    | expression                               # caseCondValue
    ;

sC_Case
    : CASE WS sC_Cond endOfStatement block?
    ;

// ELSE first, so that it is not interpreted as a variable call
sC_Cond
    : ELSE                                     # caseCondElse
    | sC_Selection (WS? ',' WS? sC_Selection)* # caseCondSelection
    ;

sendkeysStmt
    : SENDKEYS WS expression (WS? ',' WS? expression)?
    ;

setattrStmt
    : SETATTR WS expression WS? ',' WS? expression
    ;

setStmt
    : SET WS implicitCallStmt_InStmt WS? EQ WS? expression
    ;

stopStmt
    : STOP
    ;

subStmt
    : (visibility WS)? (STATIC WS)? SUB WS? ambiguousIdentifier (WS? argList)? endOfStatement block? END_SUB
    ;

timeStmt
    : TIME WS? EQ WS? expression
    ;

typeStmt
    : (visibility WS)? TYPE WS ambiguousIdentifier endOfStatement typeStmt_Element* END_TYPE
    ;

typeStmt_Element
    : ambiguousIdentifier (WS? LPAREN (WS? subscripts)? WS? RPAREN)? (WS asTypeClause)? endOfStatement
    ;

typeOfIsEzpression
    : TYPEOF WS expression (WS IS WS type_)?
    ;

unloadStmt
    : UNLOAD WS expression
    ;

unlockStmt
    : UNLOCK WS fileNumber (WS? ',' WS? expression (WS TO WS expression)?)?
    ;

// operator precedence is represented by rule order
expressiom
    : valueExpression
    | lExpression
    ;

valueExpression
    : literalExpression
    | parenthesizedExpression
    | typeOfIsExpression
    | newExpression
    | operatorExpression
    ;

lExpression
    : simpleNameExpression
    | instanceExpression
    | memberAccessExpression
    | indexExpression
    | dictionaryAccessExpression
    | withExpression
    ;

stmts
    : literal                                                                    # vsLiteral
    | implicitCallStmt_InStmt                                                    # vsICS
    | LPAREN WS? expression (WS? ',' WS? expression)* RPAREN                     # vsStruct
    | midStmt                                                                    # vsMid
    | ADDRESSOF WS? expression                                                   # vsAddressOf
    | implicitCallStmt_InStmt WS? ASSIGN WS? expression                          # vsAssign
    | expression WS? POW WS? expression                                          # vsPow
    | MINUS WS? expression                                                       # vsNegation
    | PLUS WS? expression                                                        # vsPlus
    | expression WS? (DIV | MULT) WS? expression                                 # vsDivMult
    | expression WS? MOD WS? expression                                          # vsMod
    | expression WS? (PLUS | MINUS) WS? expression                               # vsAddMinus
    | expression WS? AMPERSAND WS? expression                                    # vsAmp
    | expression WS? (IS | LIKE | GEQ | LEQ | GT | LT | NEQ | EQ) WS? expression # vsRelational
    | NOT WS? expression                                                         # vsNot
    | expression WS? AND WS? expression                                          # vsAnd
    | expression WS? OR WS? expression                                           # vsOr
    | expression WS? XOR WS? expression                                          # vsXor
    | expression WS? EQV WS? expression                                          # vsEqv
    | expression WS? IMP WS? expression                                          # vsImp
    ;

variableStmt
    : (DIM | STATIC | visibility) WS (WITHEVENTS WS)? variableListStmt
    ;

variableListStmt
    : variableSubStmt (WS? ',' WS? variableSubStmt)*
    ;

variableSubStmt
    : ambiguousIdentifier (WS? LPAREN WS? (subscripts WS?)? RPAREN WS?)? typeHint? (
        WS asTypeClause
    )?
    ;

whileWendStmt
    : WHILE WS valueStmt endOfStatement block? WEND
    ;

widthStmt
    : WIDTH WS fileNumber WS? ',' WS? valueStmt
    ;

withStmt
    : WITH WS (implicitCallStmt_InStmt | (NEW WS type_)) WS? endOfStatement block? END_WITH
    ;

writeStmt
    : WRITE WS fileNumber WS? ',' (WS? outputList)?
    ;

fileNumber
    : '#'? valueStmt
    ;

// complex call statements ----------------------------------

explicitCallStmt
    : eCS_ProcedureCall
    | eCS_MemberProcedureCall
    ;

// parantheses are required in case of args -> empty parantheses are removed
eCS_ProcedureCall
    : CALL WS ambiguousIdentifier typeHint? (WS? LPAREN WS? argsCall WS? RPAREN)? (
        WS? LPAREN subscripts RPAREN
    )*
    ;

// parantheses are required in case of args -> empty parantheses are removed
eCS_MemberProcedureCall
    : CALL WS implicitCallStmt_InStmt? '.' ambiguousIdentifier typeHint? (
        WS? LPAREN WS? argsCall WS? RPAREN
    )? (WS? LPAREN subscripts RPAREN)*
    ;

implicitCallStmt_InBlock
    : iCS_B_MemberProcedureCall
    | iCS_B_ProcedureCall
    ;

iCS_B_MemberProcedureCall
    : implicitCallStmt_InStmt? '.' ambiguousIdentifier typeHint? (WS argsCall)? dictionaryCallStmt? (
        WS? LPAREN subscripts RPAREN
    )*
    ;

// parantheses are forbidden in case of args
// variables cannot be called in blocks
// certainIdentifier instead of ambiguousIdentifier for preventing ambiguity with statement keywords
iCS_B_ProcedureCall
    : certainIdentifier (WS argsCall)? (WS? LPAREN subscripts RPAREN)*
    ;

// iCS_S_MembersCall first, so that member calls are not resolved as separate iCS_S_VariableOrProcedureCalls
implicitCallStmt_InStmt
    : iCS_S_MembersCall
    | iCS_S_VariableOrProcedureCall
    | iCS_S_ProcedureOrArrayCall
    | iCS_S_DictionaryCall
    ;

iCS_S_VariableOrProcedureCall
    : ambiguousIdentifier typeHint? dictionaryCallStmt? (WS? LPAREN subscripts RPAREN)*
    ;

iCS_S_ProcedureOrArrayCall
    : (ambiguousIdentifier | baseType) typeHint? WS? LPAREN WS? (argsCall WS?)? RPAREN dictionaryCallStmt? (
        WS? LPAREN subscripts RPAREN
    )*
    ;

iCS_S_MembersCall
    : (iCS_S_VariableOrProcedureCall | iCS_S_ProcedureOrArrayCall)? iCS_S_MemberCall+ dictionaryCallStmt? (
        WS? LPAREN subscripts RPAREN
    )*
    ;

iCS_S_MemberCall
    : WS? ('.' | '!') LINE_CONTINUATION? (
        iCS_S_VariableOrProcedureCall
        | iCS_S_ProcedureOrArrayCall
    )
    ;

iCS_S_DictionaryCall
    : dictionaryCallStmt
    ;

// atomic call statements ----------------------------------

argsCall
    : (argCall? WS? (',' | ';') WS?)* argCall (WS? (',' | ';') WS? argCall?)*
    ;

argCall
    : LPAREN? ((BYVAL | BYREF | PARAMARRAY) WS)? RPAREN? valueStmt
    ;

dictionaryCallStmt
    : '!' ambiguousIdentifier typeHint?
    ;

// atomic rules for statements

argList
    : LPAREN (WS? arg (WS? ',' WS? arg)*)? WS? RPAREN
    ;

arg
    : (OPTIONAL WS)? ((BYVAL | BYREF) WS)? (PARAMARRAY WS)? ambiguousIdentifier typeHint? (
        WS? LPAREN WS? RPAREN
    )? (WS? asTypeClause)? (WS? argDefaultValue)?
    ;

argDefaultValue
    : EQ WS? valueStmt
    ;

subscripts
    : subscript_ (WS? ',' WS? subscript_)*
    ;

subscript_
    : (valueStmt WS TO WS)? typeHint? valueStmt typeHint?
    ;

// atomic rules ----------------------------------

ambiguousIdentifier
    : (IDENTIFIER | ambiguousKeyword)+
    ;

asTypeClause
    : AS WS? (NEW WS)? type_ (WS? fieldLength)?
    ;

baseType
    : BOOLEAN
    | BYTE
    | COLLECTION
    | DATE
    | DOUBLE
    | INTEGER
    | LONG
    | SINGLE
    | STRING (WS? MULT WS? valueStmt)?
    | VARIANT
    ;

certainIdentifier
    : IDENTIFIER (ambiguousKeyword | IDENTIFIER)*
    | ambiguousKeyword (ambiguousKeyword | IDENTIFIER)+
    ;

comparisonOperator
    : LT
    | LEQ
    | GT
    | GEQ
    | EQ
    | NEQ
    | IS
    | LIKE
    ;

complexType
    : ambiguousIdentifier (('.' | '!') ambiguousIdentifier)*
    ;

fieldLength
    : MULT WS? (INTEGERLITERAL | ambiguousIdentifier)
    ;

letterrange
    : certainIdentifier (WS? MINUS WS? certainIdentifier)?
    ;

lineLabel
    : ambiguousIdentifier ':'
    ;

literal
    : HEXLITERAL
    | OCTLITERAL
    | DATELITERAL
    | DOUBLELITERAL
    | INTEGERLITERAL
    | SHORTLITERAL
    | STRINGLITERAL
    | TRUE
    | FALSE
    | NOTHING
    | NULL_
    ;

type_
    : (baseType | complexType) (WS? LPAREN WS? RPAREN)?
    ;

typeHint
    : '&'
    | '%'
    | '#'
    | '!'
    | '@'
    | '$'
    ;

visibility
    : PRIVATE
    | PUBLIC
    | FRIEND
    | GLOBAL
    ;

// ambiguous keywords
ambiguousKeyword
    : ACCESS
    | ADDRESSOF
    | ALIAS
    | AND
    | ATTRIBUTE
    | APPACTIVATE
    | APPEND
    | AS
    | BEEP
    | BEGIN
    | BINARY
    | BOOLEAN
    | BYVAL
    | BYREF
    | BYTE
    | CALL
    | CASE
    | CLASS
    | CLOSE
    | CHDIR
    | CHDRIVE
    | COLLECTION
    | CONST
    | DATABASE
    | DATE
    | DECLARE
    | DEFBOOL
    | DEFBYTE
    | DEFCUR
    | DEFDBL
    | DEFDATE
    | DEFDEC
    | DEFINT
    | DEFLNG
    | DEFOBJ
    | DEFSNG
    | DEFSTR
    | DEFVAR
    | DELETESETTING
    | DIM
    | DO
    | DOUBLE
    | EACH
    | ELSE
    | ELSEIF
    | END
    | ENUM
    | EQV
    | ERASE
    | ERROR
    | EVENT
    | FALSE
    | FILECOPY
    | FRIEND
    | FOR
    | FUNCTION
    | GET
    | GLOBAL
    | GOSUB
    | GOTO
    | IF
    | IMP
    | IMPLEMENTS
    | IN
    | INPUT
    | IS
    | INTEGER
    | KILL
    | LOAD
    | LOCK
    | LONG
    | LOOP
    | LEN
    | LET
    | LIB
    | LIKE
    | LSET
    | ME
    | MID
    | MKDIR
    | MOD
    | NAME
    | NEXT
    | NEW
    | NOT
    | NOTHING
    | NULL_
    | ON
    | OPEN
    | OPTIONAL
    | OR
    | OUTPUT
    | PARAMARRAY
    | PRESERVE
    | PRINT
    | PRIVATE
    | PUBLIC
    | PUT
    | RANDOM
    | RANDOMIZE
    | RAISEEVENT
    | READ
    | REDIM
    | REM
    | RESET
    | RESUME
    | RETURN
    | RMDIR
    | RSET
    | SAVEPICTURE
    | SAVESETTING
    | SEEK
    | SELECT
    | SENDKEYS
    | SET
    | SETATTR
    | SHARED
    | SINGLE
    | SPC
    | STATIC
    | STEP
    | STOP
    | STRING
    | SUB
    | TAB
    | TEXT
    | THEN
    | TIME
    | TO
    | TRUE
    | TYPE
    | TYPEOF
    | UNLOAD
    | UNLOCK
    | UNTIL
    | VARIANT
    | VERSION
    | WEND
    | WHILE
    | WIDTH
    | WITH
    | WITHEVENTS
    | WRITE
    | XOR
    ;

remComment
    : REMCOMMENT
    ;

comment
    : COMMENT
    ;

endOfLine
    : WS? (NEWLINE | comment | remComment) WS?
    ;

endOfStatement
    : (endOfLine | WS? COLON WS?)*
    ;

// lexer rules --------------------------------------------------------------------------------

// keywords
ACCESS
    : 'ACCESS'
    ;

ADDRESSOF
    : 'ADDRESSOF'
    ;

ALIAS
    : 'ALIAS'
    ;

AND
    : 'AND'
    ;

ATTRIBUTE
    : 'ATTRIBUTE'
    ;

APPACTIVATE
    : 'APPACTIVATE'
    ;

APPEND
    : 'APPEND'
    ;

AS
    : 'AS'
    ;

BEGIN
    : 'BEGIN'
    ;

BEEP
    : 'BEEP'
    ;

BINARY
    : 'BINARY'
    ;

BOOLEAN
    : 'BOOLEAN'
    ;

BYVAL
    : 'BYVAL'
    ;

BYREF
    : 'BYREF'
    ;

BYTE
    : 'BYTE'
    ;

CALL
    : 'CALL'
    ;

CASE
    : 'CASE'
    ;

CHDIR
    : 'CHDIR'
    ;

CHDRIVE
    : 'CHDRIVE'
    ;

CLASS
    : 'CLASS'
    ;

CLOSE
    : 'CLOSE'
    ;

COLLECTION
    : 'COLLECTION'
    ;

CONST
    : 'CONST'
    ;

DATABASE
    : 'DATABASE'
    ;

DATE
    : 'DATE'
    ;

DECLARE
    : 'DECLARE'
    ;

DEFBOOL
    : 'DEFBOOL'
    ;

DEFBYTE
    : 'DEFBYTE'
    ;

DEFDATE
    : 'DEFDATE'
    ;

DEFDBL
    : 'DEFDBL'
    ;

DEFDEC
    : 'DEFDEC'
    ;

DEFCUR
    : 'DEFCUR'
    ;

DEFINT
    : 'DEFINT'
    ;

DEFLNG
    : 'DEFLNG'
    ;

DEFOBJ
    : 'DEFOBJ'
    ;

DEFSNG
    : 'DEFSNG'
    ;

DEFSTR
    : 'DEFSTR'
    ;

DEFVAR
    : 'DEFVAR'
    ;

DELETESETTING
    : 'DELETESETTING'
    ;

DIM
    : 'DIM'
    ;

DO
    : 'DO'
    ;

DOUBLE
    : 'DOUBLE'
    ;

EACH
    : 'EACH'
    ;

ELSE
    : 'ELSE'
    ;

ELSEIF
    : 'ELSEIF'
    ;

END_ENUM
    : 'END' WS 'ENUM'
    ;

END_FUNCTION
    : 'END' WS 'FUNCTION'
    ;

END_IF
    : 'END' WS 'IF'
    ;

END_PROPERTY
    : 'END' WS 'PROPERTY'
    ;

END_SELECT
    : 'END' WS 'SELECT'
    ;

END_SUB
    : 'END' WS 'SUB'
    ;

END_TYPE
    : 'END' WS 'TYPE'
    ;

END_WITH
    : 'END' WS 'WITH'
    ;

END
    : 'END'
    ;

ENUM
    : 'ENUM'
    ;

EQV
    : 'EQV'
    ;

ERASE
    : 'ERASE'
    ;

ERROR
    : 'ERROR'
    ;

EVENT
    : 'EVENT'
    ;

EXIT_DO
    : 'EXIT' WS 'DO'
    ;

EXIT_FOR
    : 'EXIT' WS 'FOR'
    ;

EXIT_FUNCTION
    : 'EXIT' WS 'FUNCTION'
    ;

EXIT_PROPERTY
    : 'EXIT' WS 'PROPERTY'
    ;

EXIT_SUB
    : 'EXIT' WS 'SUB'
    ;

FALSE
    : 'FALSE'
    ;

FILECOPY
    : 'FILECOPY'
    ;

FRIEND
    : 'FRIEND'
    ;

FOR
    : 'FOR'
    ;

FUNCTION
    : 'FUNCTION'
    ;

GET
    : 'GET'
    ;

GLOBAL
    : 'GLOBAL'
    ;

GOSUB
    : 'GOSUB'
    ;

GOTO
    : 'GOTO'
    ;

IF
    : 'IF'
    ;

IMP
    : 'IMP'
    ;

IMPLEMENTS
    : 'IMPLEMENTS'
    ;

IN
    : 'IN'
    ;

INPUT
    : 'INPUT'
    ;

IS
    : 'IS'
    ;

INTEGER
    : 'INTEGER'
    ;

KILL
    : 'KILL'
    ;

LOAD
    : 'LOAD'
    ;

LOCK
    : 'LOCK'
    ;

LONG
    : 'LONG'
    ;

LOOP
    : 'LOOP'
    ;

LEN
    : 'LEN'
    ;

LET
    : 'LET'
    ;

LIB
    : 'LIB'
    ;

LIKE
    : 'LIKE'
    ;

LINE_INPUT
    : 'LINE' WS 'INPUT'
    ;

LOCK_READ
    : 'LOCK' WS 'READ'
    ;

LOCK_WRITE
    : 'LOCK' WS 'WRITE'
    ;

LOCK_READ_WRITE
    : 'LOCK' WS 'READ' WS 'WRITE'
    ;

LSET
    : 'LSET'
    ;

MACRO_CONST
    : '#CONST'
    ;

MACRO_IF
    : '#IF'
    ;

MACRO_ELSEIF
    : '#ELSEIF'
    ;

MACRO_ELSE
    : '#ELSE'
    ;

MACRO_END_IF
    : '#END' WS? 'IF'
    ;

ME
    : 'ME'
    ;

MID
    : 'MID'
    ;

MKDIR
    : 'MKDIR'
    ;

MOD
    : 'MOD'
    ;

NAME
    : 'NAME'
    ;

NEXT
    : 'NEXT'
    ;

NEW
    : 'NEW'
    ;

NOT
    : 'NOT'
    ;

NOTHING
    : 'NOTHING'
    ;

NULL_
    : 'NULL'
    ;

ON
    : 'ON'
    ;

ON_ERROR
    : 'ON' WS 'ERROR'
    ;

ON_LOCAL_ERROR
    : 'ON' WS 'LOCAL' WS 'ERROR'
    ;

OPEN
    : 'OPEN'
    ;

OPTIONAL
    : 'OPTIONAL'
    ;

OPTION_BASE
    : 'OPTION' WS 'BASE'
    ;

OPTION_EXPLICIT
    : 'OPTION' WS 'EXPLICIT'
    ;

OPTION_COMPARE
    : 'OPTION' WS 'COMPARE'
    ;

OPTION_PRIVATE_MODULE
    : 'OPTION' WS 'PRIVATE' WS 'MODULE'
    ;

OR
    : 'OR'
    ;

OUTPUT
    : 'OUTPUT'
    ;

PARAMARRAY
    : 'PARAMARRAY'
    ;

PRESERVE
    : 'PRESERVE'
    ;

PRINT
    : 'PRINT'
    ;

PRIVATE
    : 'PRIVATE'
    ;

PROPERTY_GET
    : 'PROPERTY' WS 'GET'
    ;

PROPERTY_LET
    : 'PROPERTY' WS 'LET'
    ;

PROPERTY_SET
    : 'PROPERTY' WS 'SET'
    ;

PTRSAFE
    : 'PTRSAFE'
    ;

PUBLIC
    : 'PUBLIC'
    ;

PUT
    : 'PUT'
    ;

RANDOM
    : 'RANDOM'
    ;

RANDOMIZE
    : 'RANDOMIZE'
    ;

RAISEEVENT
    : 'RAISEEVENT'
    ;

READ
    : 'READ'
    ;

READ_WRITE
    : 'READ' WS 'WRITE'
    ;

REDIM
    : 'REDIM'
    ;

REM
    : 'REM'
    ;

RESET
    : 'RESET'
    ;

RESUME
    : 'RESUME'
    ;

RETURN
    : 'RETURN'
    ;

RMDIR
    : 'RMDIR'
    ;

RSET
    : 'RSET'
    ;

SAVEPICTURE
    : 'SAVEPICTURE'
    ;

SAVESETTING
    : 'SAVESETTING'
    ;

SEEK
    : 'SEEK'
    ;

SELECT
    : 'SELECT'
    ;

SENDKEYS
    : 'SENDKEYS'
    ;

SET
    : 'SET'
    ;

SETATTR
    : 'SETATTR'
    ;

SHARED
    : 'SHARED'
    ;

SINGLE
    : 'SINGLE'
    ;

SPC
    : 'SPC'
    ;

STATIC
    : 'STATIC'
    ;

STEP
    : 'STEP'
    ;

STOP
    : 'STOP'
    ;

STRING
    : 'STRING'
    ;

SUB
    : 'SUB'
    ;

TAB
    : 'TAB'
    ;

TEXT
    : 'TEXT'
    ;

THEN
    : 'THEN'
    ;

TIME
    : 'TIME'
    ;

TO
    : 'TO'
    ;

TRUE
    : 'TRUE'
    ;

TYPE
    : 'TYPE'
    ;

TYPEOF
    : 'TYPEOF'
    ;

UNLOAD
    : 'UNLOAD'
    ;

UNLOCK
    : 'UNLOCK'
    ;

UNTIL
    : 'UNTIL'
    ;

VARIANT
    : 'VARIANT'
    ;

VERSION
    : 'VERSION'
    ;

WEND
    : 'WEND'
    ;

WHILE
    : 'WHILE'
    ;

WIDTH
    : 'WIDTH'
    ;

WITH
    : 'WITH'
    ;

WITHEVENTS
    : 'WITHEVENTS'
    ;

WRITE
    : 'WRITE'
    ;

XOR
    : 'XOR'
    ;

// symbols
AMPERSAND
    : '&'
    ;

ASSIGN
    : ':='
    ;

DIV
    : '\\'
    | '/'
    ;

EQ
    : '='
    ;

GEQ
    : '>='
    ;

GT
    : '>'
    ;

LEQ
    : '<='
    ;

LPAREN
    : '('
    ;

LT
    : '<'
    ;

MINUS
    : '-'
    ;

MINUS_EQ
    : '-='
    ;

MULT
    : '*'
    ;

NEQ
    : '<>'
    ;

PLUS
    : '+'
    ;

PLUS_EQ
    : '+='
    ;

POW
    : '^'
    ;

RPAREN
    : ')'
    ;

L_SQUARE_BRACKET
    : '['
    ;

R_SQUARE_BRACKET
    : ']'
    ;

// literals
fragment BLOCK
    : HEXDIGIT HEXDIGIT HEXDIGIT HEXDIGIT
    ;

GUID
    : '{' BLOCK BLOCK MINUS BLOCK MINUS BLOCK MINUS BLOCK MINUS BLOCK BLOCK BLOCK '}'
    ;

STRINGLITERAL
    : '"' (~["\r\n] | '""')* '"'
    ;

OCTLITERAL
    : '&O' [0-7]+ '&'?
    ;

HEXLITERAL
    : '&H' [0-9A-F]+ '&'?
    ;

SHORTLITERAL
    : (PLUS | MINUS)? DIGIT+
    ;

INTEGERLITERAL
    : SHORTLITERAL ('E' SHORTLITERAL)?
    ;

DOUBLELITERAL
    : (PLUS | MINUS)? DIGIT* '.' DIGIT+ ('E' SHORTLITERAL)?
    ;

DATELITERAL
    : '#' DATEORTIME '#'
    ;

fragment DATEORTIME
    : DATEVALUE WS? TIMEVALUE
    | DATEVALUE
    | TIMEVALUE
    ;

fragment DATEVALUE
    : DATEVALUEPART DATESEPARATOR DATEVALUEPART (DATESEPARATOR DATEVALUEPART)?
    ;

fragment DATEVALUEPART
    : DIGIT+
    | MONTHNAME
    ;

fragment DATESEPARATOR
    : WS? [/,-]? WS?
    ;

fragment MONTHNAME
    : ENGLISHMONTHNAME
    | ENGLISHMONTHABBREVIATION
    ;

fragment ENGLISHMONTHNAME
    : 'JANUARY'
    | 'FEBRUARY'
    | 'MARCH'
    | 'APRIL'
    | 'MAY'
    | 'JUNE | AUGUST'
    | 'SEPTEMBER'
    | 'OCTOBER'
    | 'NOVEMBER'
    | 'DECEMBER'
    ;

fragment ENGLISHMONTHABBREVIATION
    : 'JAN'
    | 'FEB'
    | 'MAR'
    | 'APR'
    | 'JUN'
    | 'JUL'
    | 'AUG'
    | 'SEP'
    | 'OCT'
    | 'NOV'
    | 'DEC'
    ;

fragment TIMEVALUE
    : DIGIT+ AMPM
    | DIGIT+ TIMESEPARATOR DIGIT+ (TIMESEPARATOR DIGIT+)? AMPM?
    ;

fragment TIMESEPARATOR
    : WS? (':' | '.') WS?
    ;

fragment AMPM
    : WS? ('AM' | 'PM' | 'A' | 'P')
    ;

// whitespace, line breaks, comments, ...
// MS-VBAL v1.7 sec 3.2.2
// *WSC underscore *WSC line-terminator
LINE_CONTINUATION
    : ' ' UNDERSCORE NEWLINE
    ;

// MS-VBAL v1.7 sec 3.2.1
NEWLINE
    : '\r\n'
    | '\r'
    | '\n'
    | '\u2028'
    | '\u2029'
    ;

REMCOMMENT
    : COLON? REM WS (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*
    ;

COMMENT
    : SINGLEQUOTE (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*
    ;

SINGLEQUOTE
    : '\''
    ;

COLON
    : ':'
    ;

UNDERSCORE
    : '_'
    ;

WS
    : ([ \t\u0019] | LINE_CONTINUATION)+
    ;

// identifier
IDENTIFIER
    : ~[\]()\r\n\t.,'"|!@#$%^&*\-+:=; ]+
    | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET
    ;

// letters
fragment LETTER
    : [A-Z_\p{L}]
    ;

fragment DIGIT
    : [0-9]
    ;

fragment HEXDIGIT
    : [A-F0-9]
    ;

fragment LETTERORDIGIT
    : [A-Z0-9_\p{L}]
    ;

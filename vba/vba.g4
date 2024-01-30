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

// module body -----
// 5.1
unrestrictedName
    : name
    | reservedIdentifier
    ;

name
    : untypedName
    | TYPEDNAME
    ;

untypedName
    : IDENTIFIER
    | FOREIGN_NAME
    ;

// expressions ----------------------------------
// 5.6
expressiom
    : valueExpression
    | lExpression
    ;

// 5.6
valueExpression
    : literalExpression
    | parenthesizedExpression
    | typeOfIsExpression
    | newExpression
    | operatorExpression
    ;

// 5.6
lExpression
    : simpleNameExpression
    | instanceExpression
    | memberAccessExpression
    | indexExpression
    | dictionaryAccessExpression
    | withExpression
    ;

// 5.6.9.3.2
additionExpression
    : expression WS? PLUS WS? expression
    ;

// 5.6.16.8
addressofExpression
    : ADDRESSOF WS? procedurePointerExpression
    ;

// 5.6.13.1
argumentExpression
    : BYVAL? expression
    | addressofExpression
    ;

// 5.6.13.1
argumentList
    : positionalOrNamedArgumentList?
    ;

positionalOrNamedArgumentList
    : (positionalArgument WS? ',')* requiredPositionalArgument
    | (positionalArgument WS? ',')* namedArgumenList
    ;

positionalArgument
    : argumentExpression?
    ;

// 5.6.16.8
procedurePointerExpression
    : simpleNameExpression
    | memberAccessExpression
    ;

requiredPositionalArgument
    : argumentExpression
    ;

namedArgumenList
    : namedArgument (',' WS? namedArgument)*
    ;

namedArgument
    : unrestrictedName WS? ASSIGN WS? argumentExpression
    ;

// 5.6.9.8.2
andOperatorExpression
    : expression WS? AND WS? wxpression
    ;

// 5.6.9.3
arithmeticOperatorExpression
    : unaryMinusExpression
    | additionExpression
    | subtractionExpression
    | multiplicationExpression
    | divisionExpression
    | integerDivisionExpression
    | moduloExpression
    | exponentiationExpression
    ;

// 5.6.9.4
concatenationOperatorExpression
    : expression WS? AMPERSAND WS? expression
    ;

// 5.6.9.3.5
divisionExpression
    : expression WS? DIV WS? expression
    ;

// 5.6.9.8.5
eqvOperatorExpression
    : expression WS? EQV WS? wxpression
    ;

// 5.6.9.3.7
exponentiationExpression
    : expression WS? POW WS? expression
    ;

// 5.6.9.5.1
equalityOperatorExpression
    : expression WS? EQ WS? expression
    ;

// 5.6.9.5.4
greaterThanOperatorExpression
    : expression WS? GT WS? expression
    ;

// 5.6.9.5.6
greaterThanEqualOperatorExpression
    : expression WS? GEQ WS? expression
    ;

// 5.6.9.8.6
impOperatorExpression
    : expression WS? IMP WS? wxpression
    ;

// 5.6.13
indexExpression
    : lExpression WS? LPAREN WS? argumentList WS? RPAREN
    ;

// 5.6.9.5.2
inequalityOperatorExpression
    : expression WS? NEQ WS? expression
    ;

// 5.6.11
instanceExpression
    : ME
    ;

// 5.6.9.3.6
integerDivisionExpression
    : expression WS? INTDIV WS? expression
    ;

// 5.6.9.7
isOperatorExpression
    : expression IS expression
    ;

// 5.6.9.5.3
lessThanOperatorExpression
    : expression WS? LT WS? expression
    ;

// 5.6.9.5.5
lessThanEqualOperatorExpression
    : expression WS? LEQ WS? expression
    ;

// 5.6.9.6
likeOperatorExpression
    : expression WS? LIKE WS? likePatternExpression
    ;

// 5.6.9.6
likePatternExpression
    : expression
    ;

// 5.6.5
literalExpression
    : INTEGER
    | FLOAT
    | DATE
    | STRING
    | (literalIdentifier typeSuffix)
    ;

// 5.6.9.8
logicalOperatorExpression
    : notOperatorExpression
    | andOperatorExpression
    | orOperatorExpression
    | xorOperatorExpression
    | impOperatorExpression
    | eqvOperatorExpression
    ;

// 5.6.12
memberAccessExpression
    : lExpression '.' WS? unrestrictedName
    | lExpression WS? LINE_CONTINUATION WS? '.' WS? unrestrictedName
    ;

// 5.6.9.3.6
moduloExpression
    : expression WS? MOD WS? expression
    ;

// 5.6.9.3.4
multiplicationExpression
    : expression WS? MULT WS? expression
    ;

// 5.6.8
newExpression
    : NEW WS? expression
    ;

// 5.6.9.8.1
notOperatorExpression
    : NOT WS? expression
    ;

// 5.6.9
operatorExpression
    : arithmeticOperatorExpression
    | concatenationOperatorExpression
    | relationalOperatorExpression
    | likeOperatorExpression
    | isOperatorExpression
    | logicalOperatorExpression
    ;

// 5.6.9.8.3
orOperatorExpression
    : expression WS? OR WS? wxpression
    ;

// 5.6.6
parenthesizedExpression
    : LPAREN WS? expression WS? RPAREN
    ;

// 5.6.9.5
relationalOperatorExpression
    : equalityOperatorExpression
    | inequalityOperatorExpression
    | lessThanOperatorExpression
    | greaterThanOperatorExpression
    | lessThanEqualOperatorExpression
    | greaterThanEqualOperatorExpression
    ;

// 5.6.10
simpleNameExpression
    : IDENTIFIER
    ;

// 5.6.9.3.3
subtractionExpression
    : expression WS? MINUS WS? expression
    ;

// 5.6.9.3.1
unaryMinusExpression
    : MINUS WS? expression
    ;

// 5.6.15
withExpression
    : withMemberAccessExpression
    | withDictionaryAccessExpression
    ;

// 5.6.15
withMemberAccessExpression
    : '.' WS? unrestrictedName
    ;

// 5.6.15
withDictionaryAccessExpression
    : '!' unrestrictedName
    ;

// 5.6.9.8.4
xorOperatorExpression
    : expression WS? XOR WS? wxpression
    ;

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
    : ambiguousIdentifier typeSuffix? (WS asTypeClause)? WS? EQ WS? expression
    ;

dateStmt
    : DATE WS? EQ WS? expression
    ;

declareStmt
    : (visibility WS)? DECLARE WS (PTRSAFE WS)? ((FUNCTION typeSuffix?) | SUB) WS ambiguousIdentifier typeSuffix? WS LIB WS STRINGLITERAL (
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
    : DELETESETTING WS expression WS? ',' WS? expression (WS? ',' WS? expression)?
    ;

doLoopStmt
    : DO endOfStatement block? LOOP
    | DO WS (WHILE | UNTIL) WS expression endOfStatement block? LOOP
    | DO endOfStatement block LOOP WS (WHILE | UNTIL) WS expression
    ;

endStmt
    : END
    ;

enumerationStmt
    : (visibility WS)? ENUM WS ambiguousIdentifier endOfStatement enumerationStmt_Constant* END_ENUM
    ;

enumerationStmt_Constant
    : ambiguousIdentifier (WS? EQ WS? expression)? endOfStatement
    ;

eraseStmt
    : ERASE WS expression (',' WS? expression)*?
    ;

errorStmt
    : ERROR WS expression
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
    : FILECOPY WS expression WS? ',' WS? expression
    ;

forEachStmt
    : FOR WS EACH WS ambiguousIdentifier typeSuffix? WS IN WS expression endOfStatement block? NEXT (
        WS ambiguousIdentifier
    )?
    ;

forNextStmt
    : FOR WS ambiguousIdentifier typeSuffix? (WS asTypeClause)? WS? EQ WS? expression WS TO WS expression (
        WS STEP WS expression
    )? endOfStatement block? NEXT (WS ambiguousIdentifier)?
    ;

functionStmt
    : (visibility WS)? (STATIC WS)? FUNCTION WS? ambiguousIdentifier typeSuffix? (WS? argList)? (
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
    : expression
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
    : (LET WS)? implicitCallStmt_InStmt WS? (EQ | PLUS_EQ | MINUS_EQ) WS? typeSuffix? expression typeSuffix?
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
    : (visibility WS)? (STATIC WS)? PROPERTY_GET WS ambiguousIdentifier typeSuffix? (WS? argList)? (
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
    : SAVEPICTURE WS expression WS? ',' WS? expression
    ;

saveSettingStmt
    : SAVESETTING WS expression WS? ',' WS? expression WS? ',' WS? expression WS? ',' WS? expression
    ;

seekStmt
    : SEEK WS fileNumber WS? ',' WS? expression
    ;

selectCaseStmt
    : SELECT WS CASE WS expression endOfStatement sC_Case* END_SELECT
    ;

sC_Selection
    : IS WS? comparisonOperator WS? expression # caseCondIs
    | expression WS TO WS expression           # caseCondTo
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

typeOfIsExpression
    : TYPEOF WS expression (WS IS WS type_)?
    ;

unloadStmt
    : UNLOAD WS expression
    ;

unlockStmt
    : UNLOCK WS fileNumber (WS? ',' WS? expression (WS TO WS expression)?)?
    ;

variableStmt
    : (DIM | STATIC | visibility) WS (WITHEVENTS WS)? variableListStmt
    ;

variableListStmt
    : variableSubStmt (WS? ',' WS? variableSubStmt)*
    ;

variableSubStmt
    : ambiguousIdentifier (WS? LPAREN WS? (subscripts WS?)? RPAREN WS?)? typeSuffix? (
        WS asTypeClause
    )?
    ;

whileWendStmt
    : WHILE WS expression endOfStatement block? WEND
    ;

widthStmt
    : WIDTH WS fileNumber WS? ',' WS? expression
    ;

withStmt
    : WITH WS (implicitCallStmt_InStmt | (NEW WS type_)) WS? endOfStatement block? END_WITH
    ;

writeStmt
    : WRITE WS fileNumber WS? ',' (WS? outputList)?
    ;

fileNumber
    : '#'? expression
    ;

// complex call statements ----------------------------------

explicitCallStmt
    : eCS_ProcedureCall
    | eCS_MemberProcedureCall
    ;

// parantheses are required in case of args -> empty parantheses are removed
eCS_ProcedureCall
    : CALL WS ambiguousIdentifier typeSuffix? (WS? LPAREN WS? argsCall WS? RPAREN)? (
        WS? LPAREN subscripts RPAREN
    )*
    ;

// parantheses are required in case of args -> empty parantheses are removed
eCS_MemberProcedureCall
    : CALL WS implicitCallStmt_InStmt? '.' ambiguousIdentifier typeSuffix? (
        WS? LPAREN WS? argsCall WS? RPAREN
    )? (WS? LPAREN subscripts RPAREN)*
    ;

implicitCallStmt_InBlock
    : iCS_B_MemberProcedureCall
    | iCS_B_ProcedureCall
    ;

iCS_B_MemberProcedureCall
    : implicitCallStmt_InStmt? '.' ambiguousIdentifier typeSuffix? (WS argsCall)? dictionaryCallStmt? (
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
    : ambiguousIdentifier typeSuffix? dictionaryCallStmt? (WS? LPAREN subscripts RPAREN)*
    ;

iCS_S_ProcedureOrArrayCall
    : (ambiguousIdentifier | baseType) typeSuffix? WS? LPAREN WS? (argsCall WS?)? RPAREN dictionaryCallStmt? (
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
    : LPAREN? ((BYVAL | BYREF | PARAMARRAY) WS)? RPAREN? expression
    ;

dictionaryCallStmt
    : '!' ambiguousIdentifier typeSuffix?
    ;

// atomic rules for statements

argList
    : LPAREN (WS? arg (WS? ',' WS? arg)*)? WS? RPAREN
    ;

arg
    : (OPTIONAL WS)? ((BYVAL | BYREF) WS)? (PARAMARRAY WS)? ambiguousIdentifier typeSuffix? (
        WS? LPAREN WS? RPAREN
    )? (WS? asTypeClause)? (WS? argDefaultValue)?
    ;

argDefaultValue
    : EQ WS? expression
    ;

subscripts
    : subscript_ (WS? ',' WS? subscript_)*
    ;

subscript_
    : (expression WS TO WS)? typeSuffix? expression typeSuffix?
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
    | STRING (WS? MULT WS? expression)?
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

// 3.3.5.3
// if changes are made here, they should be made in the
// lexer as well.
typeSuffix
    : '&'
    | '%'
    | '#'
    | '!'
    | '@'
    | '$'
    | '^'
    ;

visibility
    : PRIVATE
    | PUBLIC
    | FRIEND
    | GLOBAL
    ;

// 3.3.5.2
reservedIdentifier
    : selectKeyword
    | markerKeyword
    | operatorIdentifier
    | specialForm
    | reservedName
    | literalIdentifier
    | remKeyword
    | reservedForImplementationUse
    | futureReserved
    ;

// 3.3.5.2
remKeyword
    : REM
    ;

// 3.3.5.2
reservedForImplementationUse
    : ATTRIBUTE
    | LINEINPUT
    | VB_BASE
    | VB_CONTROL
    | VB_CREATABLE
    | VB_CUSTOMIZABLE
    | VB_DESCRIPTION
    | VB_EXPOSED
    | VB_EXT_KEY
    | VB_GLOBALNAMESPACE
    | VB_HELPID
    | VB_INVOKE_FUNC
    | VB_INVOKE_PROPERTY
    | VB_INVOKE_PROPERTYPUT
    | VB_INVOKE_PROPERTYPUTREF
    | VB_MEMBERFLAGS
    | VB_NAME
    | VB_PREDECLAREDID
    | VB_PROCDATA
    | VB_TEMPLATEDERIVED
    | VB_USERMEMID
    | VB_VARDESCRIPTION
    | VB_VARHELPID
    | VB_VARMEMBERFLAGS
    | VB_VARPROCDATA
    | VB_VARUSERMEMID
    ;

// 3.3.5.2
futureReserved
    : CDECL
    | DECIMAL
    | DEFDEC
    ;

// 3.3.5.2
selectKeyword
    : CALL
    | CASE
    | CLOSE
    | CONST
    | DECLARE
    | DEFBOOL
    | DEFBYTE
    | DEFCUR
    | DEFDATE
    | DEFDBL
    | DEFINT
    | DEFLNG
    | DEFLNGLNG
    | DEFLNGPTR
    | DEFOBJ
    | DEFSNG
    | DEFSTR
    | DEFVAR
    | DIM
    | DO
    | ELSE
    | ELSEIF
    | END
    | ENDIF
    | ENUM
    | ERASE
    | EVENT
    | EXIT
    | FOR
    | FRIEND
    | FUNCTION
    | GET
    | GLOBAL
    | GOSUB
    | GOTO
    | IF
    | IMPLEMENTS
    | INPUT
    | LET
    | LOCK
    | LOOP
    | LSET
    | NEXT
    | ON
    | OPEN
    | OPTION
    | PRINT
    | PRIVATE
    | PUBLIC
    | PUT
    | RAISEEVENT
    | REDIM
    | RESUME
    | RETURN
    | RSET
    | SEEK
    | SELECT
    | SET
    | STATIC
    | STOP
    | SUB
    | TYPE
    | UNLOCK
    | WEND
    | WHILE
    | WITH
    | WRITE
    ;

// 3.3.5.2
literalIdentifier
    : booleanLiteralIdentifier
    | objectLiteralIdentifier
    | variantLiteralIdentifier
    ;

// 3.3.5.2
booleanLiteralIdentifier
    : TRUE
    | FALSE
    ;

// 3.3.5.2
objectLiteralIdentifier
    : NOTHING
    ;

// 3.3.5.2
variantLiteralIdentifier
    : EMPTY
    | NULL
    ;

// 3.3.5.2
markerKeyword
    : ANY
    | AS
    | BYREF
    | BYVAL
    | CASE
    | EACH
    | ELSE
    | IN
    | NEW
    | SHARED
    | UNTIL
    | WITHEVENTS
    | WRITE
    | OPTIONAL
    | PARAMARRAY
    | PRESERVE
    | SPC
    | TAB
    | THEN
    | TO
    ;

// 3.3.5.2
operatorIdentifier
    : ADDRESSOF
    | AND
    | EQV
    | IMP
    | IS
    | LIKE
    | NEW
    | MOD
    | NOT
    | OR
    | TYPEOF
    | XOR
    ;

// 3.3.5.2
reservedName
    : ABS
    | CBOOL
    | CBYTE
    | CCUR
    | CDATE
    | CDBL
    | CDEC
    | CINT
    | CLNG
    | CLNGLNG
    | CLNGPTR
    | CSNG
    | CSTR
    | CVAR
    | CVERR
    | DATE
    | DEBUG
    | DOEVENTS
    | FIX
    | INT
    | LEN
    | LENB
    | ME
    | PSET
    | SCALE
    | SGN
    | STRING
    ;

// 3.3.5.2
specialForm
    : ARRAY
    | CIRCLE
    | INPUT
    | INPUTB
    | LBOUND
    | SCALE
    | UBOUND
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

INTDIV
    : '\\'
    ;

DIV
    : '/'
    ;

EQ
    : '='
    ;

GEQ
    : '>='
    | '=>'
    ;

GT
    : '>'
    ;

LEQ
    : '<='
    | '=<'
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
    | '><'
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

TYPED_NAME
    : IDENTIFIER [%$^!#@$]
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
// 3.3.5.3
FOREIGN_NAME
    : L_SQUARE_BRACKET (~[\]\r\n])+ R_SQUARE_BRACKET
    ;

IDENTIFIER
    : ~[\]()\r\n\t.,'"|!@#$%^&*\-+:=; ]+
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
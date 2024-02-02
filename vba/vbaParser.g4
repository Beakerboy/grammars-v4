/*
* Visual Basic 7.1 Grammar for ANTLR4
*
* This is an approximate grammar for Visual Basic 6.0, derived
* from the Visual Basic 6.0 language reference
* https://msopenspecs.azureedge.net/files/MS-VBAL/%5bMS-VBAL%5d.pdf
* and tested against MSDN VB6 statement examples as well as several Visual
* Basic 6.0 code repositories.
*/

// $antlr-format alignTrailingComments true, columnLimit 150, minEmptyLines 1, maxEmptyLinesToKeep 1, reflowComments false, useTab false
// $antlr-format allowShortRulesOnASingleLine false, allowShortBlocksOnASingleLine true, alignSemicolons hanging, alignColons hanging

parser grammar vbaParser;


options {
    caseInsensitive = true;
    tokenVocab = vbaLexer;
}

// Contexts not listed in the specification
// Everything until section 5.1 is typically machine generated code.
startRule
    : module EOF
    ;

module
    : WS? endOfLine* (
          proceduralModule
        | classFileHeader endOfLine+ classModule
      ) endOfLine* WS?
    ;

classFileHeader
    : classVersionIdentification endOfLine+ classBeginBlock endOfLine+
    ;

classVersionIdentification
    : VERSION WS DOUBLELITERAL (WS CLASS)?
    ;

classBeginBlock
    : BEGIN (WS GUID WS ambiguousIdentifier)? endOfLine* beginBlockConfigElement+ END
    ;

beginBlockConfigElement
    : ambiguousIdentifier WS? EQ WS? literalExpression (COLON literalExpression)? endOfLine*
    ;

// 4.2 Modules
proceduralModule
    : proceduralModuleHeader endOfLine* proceduralModuleBody
    ;
classModule
    : classModuleHeader endOfLine* classModuleBody
    ;

// Compare STRINGLITERAL to quoted-identifier
proceduralModuleHeader
    : ATTRIBUTE WS? VB_NAME WS? EQ WS? STRINGLITERAL endOfLine
    ;
classModuleHeader: classAttr+;
classAttr
    : ATTRIBUTE WS? VB_NAME WS? EQ WS? STRINGLITERAL endOfLine
    | ATTRIBUTE WS? VB_GLOBALNAMESPACE WS? EQ WS? FALSE endOfLine
    | ATTRIBUTE WS? VB_CREATABLE WS? EQ WS? FALSE endOfLine
    | ATTRIBUTE WS? VB_PREDECLAREDID WS? EQ WS? booleanLiteralIdentifier endOfLine
    | ATTRIBUTE WS? VB_EXPOSED WS? EQ WS? booleanLiteralIdentifier endOfLine
    | ATTRIBUTE WS? VB_CUSTOMIZABLE WS? EQ WS? booleanLiteralIdentifier endOfLine
    ;

// 5.1 Module Body Structure
// Everthing from here down is user generated code.
proceduralModuleBody: proceduralModuleDeclarationSection? proceduralModuleCode;
classModuleBody: classModuleDeclarationSection? classModuleCode;

// 5.2 Module Declaration Section Structure
proceduralModuleDeclarationSection
    : ((proceduralModuleDirectiveElement endOfLine+)* defDirective)? (proceduralModuleDeclarationElement endOfLine)*
    ;
classModuleDeclarationSection
    : ((classModuleDirectiveElement endOfLine+)* defDirective)? (classModuleDeclarationElement endOfLine)*
    ;
proceduralModuleDirectiveElement
    : commonOptionDirective
    | optionPrivateDirective
    | defDirective
    ;
proceduralModuleDeclarationElement
    : commonModuleDeclarationElement
    | globalVariableDeclaration
    | publicConstDeclaration
    | publicTypeDeclaration
    | publicExternalProcedureDeclaration
    | globalEnumDeclaration
    | commonOptionDirective
    | optionPrivateDirective
    ;
classModuleDirectiveElement
    : commonOptionDirective
    | defDirective
    | implementsDirective
    ;
classModuleDeclarationElement
    : commonModuleDeclarationElement
    | eventDeclaration
    | commonoptionDirective
    | implementsDirective
    ;

// 5.2.1 Option Directives
commonOptionDirective
    : optionCompareDirective
    | optionBaseDirective
    | optionExplicitDirective
    | remStatement
    ;

// 5.2.1.1 Option Compare Directive
optionCompareDirective: OPTION wsc COMPARE wsc (BINARY | TEXT);

// 5.2.1.2 Option Base Directive
// INTEGER or SHORT?
optionBaseDirective: OPTION wsc BASE wsc SHORTLITERAL;

// 5.2.1.3 Option Explicit Directive
optionExplicitDirective: OPTION wsc EXPLICIT;

// 5.2.1.4 Option Private Directive
optionPrivateDirective: OPTION wsc PRIVATE wsc MODULE;

// 5.2.2 Implicit Definition Directives
defDirective: defType WS letterSpec (WS ',' WS letterSpec)*;
letterSpec
    : singleLetter
    | universalLetterRange
    | letterRange
    ;
singleLetter: ambiguousIdentifier;
universalLetterRange: upperCaseA WS '-' WS upperCaseZ;
upperCaseA: ambiguousIdentifier;
upperCaseZ: ambiguousIdentifier;
letterRange: firstLetter WS '-' WS lastLetter;
firstLetter: ambiguousIdentifier;
lastLetter: ambiguousIdentifier;
defType
    : DEFBOOL
    | DEFBYTE
    | DEFCUR
    | DEFDATE
    | DEFDBL
    | DEFINT
    | DEFINT
    | DEFLNG
    | DEFLNGLNG
    | DEGLNGPTR
    | DEFOBJ
    | DEFSNG
    | DEFSTR
    | DEFVAR
    ;

// 5.2.3 Module Declarations
commonModuleDeclarationElement
    : moduleVariableDeclaration
    | privateConstDeclaration
    | privateTypeDeclaration
    | enumDeclaration
    | privateExternalProcedureDeclaration
    ;

// 5.2.3.1 Module Variable Declaration Lists
globalVariableDeclaration: GLOBAL WS variableDeclarationList;
publicVariableDecalation: PUBLIC (WS SHARED)? WS moduleVariableDeclarationList;
privateVariableDeclaration: (PRIVATE | DIM) (wsc SHARED)? moduleVariableDeclarationList;
moduleVariableDeclarationList: (witheventsVariableDcl | variableDcl) (wsc? ',' wsc? (witheventsVariableDcl | variableDcl))*;
variableDeclarationList: variableDcl wsc (wsc? ',' wsc? variableDcl)*;

// 5.2.3.1.1 Variable Declarations
variableDcl
    : typedVariableDcl
    | untypedVariableDcl
    ;
typedVariableDcl: typedName arrayDim?;
untypedVariableDcl: ambiguousIdentifier (arrayClause | asClause)?;
arrayClause: arrayDim wsc? asClause;
asClause
    : asAutoObject
    | asType
    ;

// 5.2.3.1.2 WithEvents Variable Declarations
witheventsVariableDcl: WITHEVENTS wsc ambiguousIdentifier wsc AS wsc? classTypeName;
classTypeName: definedTypeExpression;

// 5.2.3.1.3 Array Dimensions and Bounds
arrayDim: '(' wsc? boundsList? wsc? ')';
boundsList: dimSpec (wsc',' wsc dimSpec)*;
dimSpec: lowerBound? wsc? upperBound;
lowerBound: constantExpression TO;
upperBound: constantExpression;

// 5.2.3.1.4 Variable Type Declarations
asAutoObject: AS WS NEW WS classTypeName;
asType: AS WS typeSpec;
typeSpec
    : fixedLengthStringSpec
    | typeExpression
    ;
fixedLengthStringSpe: STRING WS '*' WS stringLength;
stringLength
    : SHORTLITERAL
    | constantName
    ;
constantName: simpleNameExpression;

// 5.2.3.2 Const Declarations
publicConstDeclaration: (GLOBAL | PUBLIC) wsc module_const_declaration;
privateConstDeclaration: PRIVATE wsc module_const_declaration;
moduleConstDeclaration: constDeclaration;
constDeclaration: CONST wsc constItemList;
constItemList: constItem (wsc? ',' wsc? constItem)*;
constItem
    : typedNameConstItem
    | untypedNameConstItem
    ;
typedNameConstItem: typedName wsc? EQ wsc? constantExpression;
untypedNameConstItem: ambiguousIdentifier (wsc constAsClause)? wsc? EQ wsc? constantExpression;
constAsClause: builtinType;

// 5.2.3.3 User Defined Type Declarations
publicTypeDeclaration: (GLOBAL | PUBLIC) wsc udrDeclaration;
privateTypeDeclaration: PRIVATE wsc udtDeclaration;
udtDeclaration: TYPE wsc untypedName endOfStatement+ udtMemberList endOfStatement+ END wsc TYPE
udtMemberList: udtElement wsc (endOfStatement udtElement)*
udtElement
    : remStatement
    | udtMember
    ;
udtMember:
    : reservedNameMemberDcl
    | untypedNameMemberDcl
    ;
untypedNameMemberDcl: ambiguousIdentifier optionalArrayClause;
reservedNameMemberDcl: reservedMemberName wsc asClause;
reservedMemberName
    : statementKeyword
    | markerKeyword
    | operatorIdentifier
    | specialForm
    | reservedName
    | literalIdentifier
    | reservedForImplementationUse
    | futureReserved
    ;

// 5.2.3.4 Enum Declarations
globalEnumDeclaration: GLOBAL wsc  enumDeclaration;
publicEnumDeclaration: (PUBLIC wsc)? enumDeclaration;
privateEnumDeclaration: PRIVATE wsc enumDeclaration;
enumDeclaration: ENUM wsc untypedName EOS memberList EOS END wsc ENUM ;
enumMemberList: enumElement (endOfStatement enumElement)*;
enumElement
    : remStatement
    | enumMember
    ;
enumMember: untyped-name (wsc? EQ wsc? constantExpression)?;

// 5.2.3.5 External Procedure Declaration
public-external-procedure-declaration: (PUBLIC wsc)? externalProcDcl;
private-external-procedure-declaration: PRIVATE externalProcDcl;
externalProcDcl: DECLARE wsc (PTRSAFE wsc)? (externalSub | externalFunction);
externalSub: SUB subroutine-name lib-info [procedure-parameters];
externalFunction: FUNCTION functionName libInfo procedureParameters? functionType?;
libInfo: libClause (wsc aliasClause)?;
libClause: LIB wsc STRING;
aliasClause: ALIAS wsc STRING;

// 5.3
commonModuleCodeElement
    : remStatement
    | procedureDeclaration
    ;

procedureDeclaration
    : functionStmt
    | propertyGetStmt
    | propertySetStmt
    | propertyLetStmt
    | subroutineDeclaration
    | macroStmt
    ;
// block ----------------------------------

// body ------------------------------

implementsDirective
    : IMPLEMENTS WS ambiguousIdentifier
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



proceduralModuleCode
    : proceduralModuleCodeElement (endOfLine+ proceduralModuleCodeElement)* endOfLine*
    ;

classModuleCode
    : classModuleCodeElement (endOfLine+ classModuleCodeElement)* endOfLine*
    ;

proceduralModuleCodeElement
    : commonModuleCodeElement
    ;

classModuleCodeElement
    : commonModuleCodeElement
    | implementsDirective
    ;

// 5.4
procedureBody
    : statementBlock
    ;

statementBlock
    : (blockStmt endOfStatement)*
    ;
block
    : blockStmt (endOfStatement blockStmt)* endOfStatement
    ;

blockStmt
    : statementLabelDefinition
    | remStatement
    | statement
    ;

statement
    : controlStatement
    | dataManipulationStatement
    | errorHandingStatement
    | fileStatement
    ;
    
// 5.4.1.1
statementLabelDefinition
    : (identifierStatementLabel | lineNumberLabel) ':'
    ;

lineNumberLabel
    : (INTEGERLITERAL | SHORTLITERAL)
    ;

identifierStatementLabel
    : ambiguousIdentifier
    ;
    
// 5.4.2
controlStatement
    : ifThenElseStmt
    | controlStatementExceptMultilineIf
    ;

controlStatementExceptMultilineIf
    : callStatement
    | whileStatement
    | forStatement
    | exitForStatement
    | doStatement
    | exitDoStatement
    | singleLineIfStatement
    | selectCaseStatement
    | stopStatement
    | gotoStatement
    | onGotoStatement
    | gosubStatement
    | returnStatement
    | onGosubStatement
    | forEachStatement
    | exitSubStatement
    | exitFunctionStatement
    | exitPropertyStatement
    | raiseeventStatement
    | withStatement
    ;
extra
    : appactivateStmt
    | beepStmt
    | chdirStmt
    | chdriveStmt
    | constStmt
    | dateStmt
    | deleteSettingStmt
    | endStmt
    | eraseStmt
    | explicitCallStmt
    | filecopyStmt
    | forNextStmt
    | implementsStmt
    | killStmt
    | loadStmt
    | macroStmt
    | mkdirStmt
    | nameStmt
    | onErrorStmt
    | randomizeStmt
    | resumeStmt
    | rmdirStmt
    | savepictureStmt
    | saveSettingStmt
    | sendkeysStmt
    | setattrStmt
    | timeStmt
    | unloadStmt
    | variableStmt
    | expression
    ;


// 5.4.2.1
callStatement
    : CALL WS (simpleNameExpression
        | memberAccessExpression
        | indexExpression
        | withExpression)
    | (simpleNameExpression
        | memberAccessExpression
        | withExpression) argumentList
    ;

// 5.4.3
dataManipulationStatement
    : localVariableDeclaration
    | staticVariableDeclaration
    | localConstDeclaration
    | redimStatement
    | midStatement
    | rsetStatement
    | lsetStatement
    | letStatement
    | setStatement
    ;

// 5.4.3.1
localVariableDeclaration: DIM wsc? SHARED? wsc? variableDeclarationList;
staticVariableDeclaration: STATIC wsc variableDeclarationList;

// 5.4.5
fileStatement
    : openStatement
    | close Statement
    | seekStatement
    | lockStatement
    | unlockStatement
    | lineInputStatement
    | widthStatement
    | writeStatement
    | inputStatement
    | putStatement
    | getStatement
    ;

// 5.4.5.1
openStatement
    : OPEN wcs? pathName wsc? modeClause? wsc? accessClause? wsc? lock? wsc? AS wsc? fileNumber wsc? lenClause?
    ;
pathName: expression;
modeClause: FOR wsc? modeOpt;
modeOpt
    : APPEND
    | BINARY
    | INPUT
    | OUTPUT
    | RANDOM
    ;
accessClause: ACCESS access;
access
    : READ
    | WRITE
    | READ wsc WRITE
    ;
lock
    : SHARED
    | LOCK wsc READ
    | LOCK wsc WRITE
    | LOCK wsc READ wsc WRITE
    ;
lenClause: LEN wsc EQ wsc recLength;
recLength: expression;

// 5.4.5.1.1
fileNumber
    : markedFileNumber
    | unmarkedFileNumber
    ;
markedFileNumber: '#' expression;
unmarkedFileNumber: expression;

// 5.4.5.2
closeStatement
    : RESET
    | CLOSE wsc? fileNumberList?
    ;
fileNumberList: fileNumber (wsc? ',' wsc? fileNumber)*;

// 5.4.5.3
seekStatement: SEEK wsc fileNumber wsc? ',' wsc? position;
position: expression;

// 5.4.5.4
lockStatement: LOCK wsc fileNumber (wsc? ',' wsc? recordRange);
recordRange
    : startRecordNumber
    | startRecordNumber? wsc TO wsc endRecordNumber
    ;
startRecordNumber: expression;
endRecordNumber: expression;

// 5.4.5.5
unlockStatement: UNLOCK wsc fileNumber (wsc? ',' wsc? recordRange)?;

// 5.4.5.6
lineInputStatement: LINE wsc INPUT wsc markedFileNumber wsc? ',' wsc? variableName;
variableName: variableExpression;

// 5.4.5.7
widthStatement: WIDTH wsc markedFileNumber wsc? ',' wsc? lineWidth;
lineWidth: expression;

// 5.4.5.8
printStatement: PRINT wsc markedFileNumber wsc? ',' wsc? outputList?;

// 5.4.5.8.1
outputList: outputItem*;
outputItem: outputClause? charPosition?;
outputClause: spcClause | tabClause| outputExpression;
charPosition: ';' | ',';
outputExpression: expression;
spcClause: SPC wsc '(' wsc? spcNumber wsc? ')';
spcNumber: expression;
tabClause: TAB wsc '(' wsc? tabNumber wsc? ')';
tabNumbeer: expression;

// 5.4.5.9
writeStatement: WRITE wsc markedFileNumber wsc? ',' wsc? outputList?;

// 5.4.5.10
inputStatement: INPUT wsc markedFileNumber wsc? ',' wsc? inputList;
inputList: inputVariable (wsc? ',' wsc? inputVariable)*;
inputVariable: boundVariableExpression;

//5.4.5.11
putStatement: PUT wsc fileNumber wsc? ',' wsc?recordNumber? wsc? ',' data;
recordNumber: expression;
data: expression;

// 5.4.5.12
getStatement: GET wsc fileNumber wsc? ',' wsc? recordNumber? wsc? ',' ws? variable;
variable: variableExpression;

// expressions----------------------------------
// 5.6
// Modifying the order will affect the order of operations
// valueExpression must be rolled up into expression due to mutual left recursion
expression
    : literalExpression
    | parenthesizedExpression
    | typeOfStmt
    | newExpress
    | midStatement
    | ADDRESSOF wsc? expression
    | implicitCallStmt_InStmt wsc? ASSIGN wsc? expression
    | expression wsc? POW wsc? expression
    | unaryMinusExpression
    | expression wsc? (DIV | MULT) wsc? expression
    | expression wsc? MOD wsc? expression
    | expression wsc? (PLUS | MINUS) wsc? expression
    | expression wsc? AMPERSAND wsc? expression
    | expression wsc? (IS | LIKE | GEQ | LEQ | GT | LT | NEQ | EQ) wsc? expression
    | notOperatorExpression
    | expression wsc? (AND | OR | XOR | EQV | IMP) wsc? expression
    | lExpression
    ;

// Several of the lExpression rules are rolled up due to Mutual Left Recursion
// Many are also listed separately due to their specific use elsewhere.
lExpression
    : simpleNameExpression
    | instanceExpression
// memberAccessExpression
    | lExpression '.' WS? unrestrictedName
    | lExpression WS? LINE_CONTINUATION WS?'.' WS? unrestrictedName
//
    | lExpression WS? '(' WS? argumentList WS ')'
    ;
    
// 5.6.5
// check on hex and oct
// check definition of integer and float
literalExpression
    : HEXLITERAL
    | OCTLITERAL
    | DATELITERAL
    | DOUBLELITERAL
    | INTEGERLITERAL
    | SHORTLITERAL
    | STRINGLITERAL
    | literalIdentifier typeSuffix?
    ;

// 5.6.6
parenthesizedExpression
    : LPAREN wsc? expression wsc? RPAREN
    ;

// 5.6.8
// The name 'newExpression' fails under the Go language
newExpress
    : NEW wsc? expression
    ;

// 5.6.9.8.1
notOperatorExpression
    : NOT wsc? expression
    ;

// 5.6.9.3.1
unaryMinusExpression
    : MINUS wsc? expression
    ;

// 5.6.10
simpleNameExpression
    : name
    ;

// 5.6.11
instanceExpression
    : ME
    ;

// 5.6.12
// This expression is also rolled into lExpression
// changes here must be duplicated there
memberAccessExpression
    : lExpression '.' WS? unrestrictedName
    | lExpression WS? LINE_CONTINUATION WS?'.' WS? unrestrictedName
    ;

// 5.6.13
indexExpression
    : lExpression WS? '(' WS? argumentList WS ')'
    ;

argumentList
    : positionalOrNamedArgumentList?
    ;

positionalOrNamedArgumentList
    : (positionalArgument WS? ',')* requiredPositionalArgument
    | (positionalArgument WS? ',')* namedArgumentList
    ;

positionalArgument
    : argumentExpression?
    ;

requiredPositionalArgument
    : argumentExpression
    ;

namedArgumentList
    : namedArgument (wsc? ',' wsc? namedArgument)*
    ;

namedArgument
    : unrestrictedName wsc? ASSIGN wsc? argumentExpression
    ;

argumentExpression
    : BYVAL? wsc? expression
    | addressofExpression
    ;

// 5.6.16.1
// This could be made more complicated for accuracy
constantExpression: expression;

// 5.6.16.5
variableExpression: lExpression;

// 5.6.16.6
boundVariableExpression: lExpression;

// 5.6.16.7
typeExpression
    : builtinType
    | definedTypeExpression
    ;
definedTypeExpression
    : simpleNameExpression
    | memberAccessExpression
    ;

// 5.6.16.8
addressofExpression
    : ADDRESSOF procedurePointerExpression
    ;

procedurePointerExpression
    : simpleNameExpression
    | memberAccessExpression
    ;

variableStmt
    : (DIM | STATIC | visibility) WS (WITHEVENTS WS)? variableListStmt
    ;

variableListStmt
    : variableSubStmt (WS? ',' WS? variableSubStmt)*
    ;

variableSubStmt
    : ambiguousIdentifier (WS? LPAREN WS? (subscripts WS?)? RPAREN WS?)? typeSuffix? (
        WS asType
    )?
    ;

whileWendStmt
    : WHILE WS expression endOfStatement block? WEND
    ;

widthStmt
    : WIDTH WS fileNumber WS? ',' WS? expression
    ;

withStmt
    : WITH WS (implicitCallStmt_InStmt | (NEW WS type_)) endOfStatement block? END_WITH
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
    : ambiguousIdentifier typeSuffix? (WS asType)? WS? EQ WS? expression
    ;

dateStmt
    : DATE WS? EQ WS? expression
    ;

declareStmt
    : (visibility WS)? DECLARE WS (PTRSAFE WS)? ((FUNCTION typeSuffix?) | SUB) WS ambiguousIdentifier typeSuffix? WS LIB WS STRINGLITERAL (
        WS ALIAS WS STRINGLITERAL
    )? (WS? argList)? (WS asType)?
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

doStatement
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

errorHandingStatement
    : ERROR WS expression
    ;

eventStmt
    : (visibility WS)? EVENT WS ambiguousIdentifier WS? argList
    ;

exitDoStatement
    : EXIT wsc DO
    ;
exitForStatement
    : EXIT wsc FOR
    ;
exitSubStatement
    : EXIT wsc SUB
    ;
exitPropertyStatement
    : EXIT wsc PROPERTY
    ;
exitFunctionStatement
    : EXIT wsc FUNCTION
    ;

filecopyStmt
    : FILECOPY WS expression WS? ',' WS? expression
    ;

forEachStatement
    : FOR WS EACH WS ambiguousIdentifier typeSuffix? WS IN WS expression endOfStatement block? NEXT (
        WS ambiguousIdentifier
    )?
    ;

forNextStmt
    : FOR WS ambiguousIdentifier typeSuffix? (WS asType)? WS? EQ WS? expression WS TO WS expression (
        WS STEP WS expression
    )? endOfStatement block? NEXT (WS ambiguousIdentifier)?
    ;

functionStmt
    : (visibility WS)? (STATIC WS)? FUNCTION WS? ambiguousIdentifier typeSuffix? (WS? argList)? (
        WS? asType
    )? endOfStatement block? END wsc FUNCTION
    ;

gosubStatement
    : GOSUB WS expression
    ;

gotoStatement
    : GOTO WS expression
    ;

ifThenElseStmt
    : IF WS ifConditionStmt WS THEN WS blockStmt (WS ELSE WS blockStmt)? # inlineIfThenElse
    | ifBlockStmt ifElseIfBlockStmt* ifElseBlockStmt? END wsc IF             # blockIfThenElse
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

letStatement
    : (LET WS)? implicitCallStmt_InStmt WS? (EQ | PLUS_EQ | MINUS_EQ) WS? typeSuffix? expression typeSuffix?
    ;

lineInputStmt
    : LINE_INPUT WS fileNumber WS? ',' WS? expression
    ;


loadStmt
    : LOAD WS expression
    ;

lockStmt
    : LOCK WS expression (WS? ',' WS? expression (WS TO WS expression)?)?
    ;

lsetStatement
    : LSET WS implicitCallStmt_InStmt WS? EQ WS? expression
    ;

macroConstStmt
    : MACRO_CONST WS? ambiguousIdentifier WS? EQ WS? expression
    ;

macroIfThenElseStmt
    : macroIfBlockStmt macroElseIfBlockStmt* macroElseBlockStmt? MACRO_END_IF
    ;

macroIfBlockStmt
    : MACRO_IF WS? ifConditionStmt WS THEN endOfStatement (moduleDeclarations | proceduralModuleBody | classModuleBody | block)*
    ;

macroElseIfBlockStmt
    : MACRO_ELSEIF WS? ifConditionStmt WS THEN endOfStatement (
        moduleDeclarations
        | proceduralModuleBody | classModuleBody
        | block
    )*
    ;

macroElseBlockStmt
    : MACRO_ELSE endOfStatement (moduleDeclarations | proceduralModuleBody | classModuleBody | block)*
    ;

midStatement
    : MID WS? LPAREN WS? argsCall WS? RPAREN
    ;

mkdirStmt
    : MKDIR WS expression
    ;

nameStmt
    : NAME WS expression WS AS WS expression
    ;

onErrorStmt
    : (ON_ERROR | ON_LOCAL_ERROR) WS (GOTO WS expression | RESUME WS NEXT)
    ;

onGotoStatement
    : ON WS expression WS GOTO WS expression (WS? ',' WS? expression)*
    ;

onGosubStatement
    : ON WS expression WS GOSUB WS expression (WS? ',' WS? expression)*
    ;

openStmt
    : OPEN WS expression WS FOR WS (APPEND | BINARY | INPUT | OUTPUT | RANDOM) (
        WS ACCESS WS (READ | WRITE | READ_WRITE)
    )? (WS (SHARED | LOCK_READ | LOCK_WRITE | LOCK_READ_WRITE))? WS AS WS fileNumber (
        WS LEN WS? EQ WS? expression
    )?
    ;

printStmt
    : PRINT WS fileNumber WS? ',' (WS? outputList)?
    ;

propertyGetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_GET WS ambiguousIdentifier typeSuffix? (WS? argList)? (
        WS asType
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

redimStatement
    : REDIM WS (PRESERVE WS)? redimSubStmt (WS? ',' WS? redimSubStmt)*
    ;

redimSubStmt
    : implicitCallStmt_InStmt WS? LPAREN WS? subscripts WS? RPAREN (WS asType)?
    ;

resetStatement
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

rsetStatement
    : RSET WS implicitCallStmt_InStmt WS? EQ WS? expression
    ;

savepictureStmt
    : SAVEPICTURE WS expression WS? ',' WS? expression
    ;

saveSettingStmt
    : SAVESETTING WS expression WS? ',' WS? expression WS? ',' WS? expression WS? ',' WS? expression
    ;

selectCaseStatement
    : SELECT WS CASE WS expression endOfStatement sC_Case* END_SELECT
    ;

sC_Selection
    : IS WS? comparisonOperator WS? expression # caseCondIs
    | expression WS TO WS expression            # caseCondTo
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

setStatement
    : SET WS implicitCallStmt_InStmt WS? EQ WS? expression
    ;

stopStatement
    : STOP
    ;

subroutineDeclaration
    : (visibility WS)? (STATIC WS)? SUB WS? ambiguousIdentifier (WS? argList)? endOfStatement block? END wsc SUB procedureTail?
    ;

procedureTail
    : WS? NEWLINE
    | COMMENT
    | REMCOMMENT
    ;
timeStmt
    : TIME WS? EQ WS? expression
    ;

typeStmt
    : (visibility WS)? TYPE WS ambiguousIdentifier endOfStatement typeStmt_Element* END_TYPE
    ;

typeStmt_Element
    : ambiguousIdentifier (WS? LPAREN (WS? subscripts)? WS? RPAREN)? (WS asType)? endOfStatement
    ;

typeOfStmt
    : TYPEOF WS expression (WS IS WS type_)?
    ;

unloadStmt
    : UNLOAD WS expression
    ;

unlockStmt
    : UNLOCK WS fileNumber (WS? ',' WS? expression (WS TO WS expression)?)?
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
    : (ambiguousIdentifier | reservedTypeIdentifier) typeSuffix? WS? LPAREN WS? (argsCall WS?)? RPAREN dictionaryCallStmt? (
        WS? LPAREN subscripts RPAREN
    )*
    ;

iCS_S_MembersCall
    : (iCS_S_VariableOrProcedureCall | iCS_S_ProcedureOrArrayCall)? iCS_S_MemberCall+ dictionaryCallStmt? (
        WS? LPAREN subscripts RPAREN
    )*
    ;

iCS_S_MemberCall
    : LINE_CONTINUATION? WS? ('.' | '!') LINE_CONTINUATION? WS? (
        iCS_S_VariableOrProcedureCall
        | iCS_S_ProcedureOrArrayCall
    )
    ;

iCS_S_DictionaryCall
    : dictionaryCallStmt
    ;

// atomic call statements ----------------------------------

argsCall
    : (argCall? wsc? (',' | ';') wsc?)* argCall (wsc? (',' | ';') wsc? argCall?)*
    ;

argCall
    : LPAREN? ((BYVAL | BYREF | PARAMARRAY) WS)? RPAREN? expression
    ;

dictionaryCallStmt
    : '!' ambiguousIdentifier typeSuffix?
    ;

// atomic rules for statements

argList
    : LPAREN (wsc? arg (wsc? ',' wsc? arg)*)? wsc? RPAREN
    ;

arg
    : (OPTIONAL WS)? ((BYVAL | BYREF) WS)? (PARAMARRAY WS)? ambiguousIdentifier typeSuffix? (
        WS? LPAREN WS? RPAREN
    )? (WS? asType)? (WS? argDefaultValue)?
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

unrestrictedName
    : reservedIdentifier
    | ambiguousIdentifier
    ;
// Known as IDENTIFIER in MS-VBAL
ambiguousIdentifier
    : IDENTIFIER
    | ambiguousKeyword
    ;

builtinType
    : reservedTypeIdentifier
    | '[' reservedTypeIdentifier ']'
    | OBJECT
    | '[' OBJECT ']'
    ;
name
    : untypedName
    | typedName
    ;

untypedName
    : ambiguousIdentifier
    ;

typedName
    : ambiguousIdentifier typeSuffix
    ;
    

booleanLiteralIdentifier
    : TRUE
    | FALSE
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

futureReserved
    : CDECL
    | DECIMAL
    | DEFDEC
    ;

letterrange
    : certainIdentifier (WS? MINUS WS? certainIdentifier)?
    ;


literalIdentifier
    : booleanLiteralIdentifier
    | objectLiteralIdentifier
    | variantLiteralIdentifier
    ;

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

objectLiteralIdentifier
    : NOTHING
    ;

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

reservedTypeIdentifier
    : BOOLEAN
    | BYTE
    | CURRENCY
    | DATE
    | DOUBLE
    | INTEGER
    | LONG
    | LONGLONG
    | LONGPTR
    | SINGLE
    | STRING
    | VARIANT
    ;

specialForm
    : ARRAY
    | CIRCLE
    | INPUT
    | INPUTB
    | LBOUND
    | SCALE
    | UBOUND
    ;

statementKeyword
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

type_
    : (reservedTypeIdentifier | complexType) (WS? LPAREN WS? RPAREN)?
    ;

typeSuffix
    : '&'
    | '%'
    | '#'
    | '!'
    | '@'
    | '$'
    | '^'
    ;

variantLiteralIdentifier
    : EMPTY
    | NULL_
    ;

visibility
    : PRIVATE
    | PUBLIC
    | FRIEND
    | GLOBAL
    ;

// should this include reservedTypeIdentifier?
reservedIdentifier
    : statementKeyword
    | markerKeyword
    | operatorIdentifier
    | specialForm
    | reservedName
    | literalIdentifier
    | remKeyword
    | reservedForImplementationUse
    | futureReserved
    ;

// lexer keywords not in the reservedIdentifier set
// should this include reservedTypeIdentifier?
ambiguousKeyword
    : ACCESS
    | ALIAS
    | APPACTIVATE
    | APPEND
    | BEEP
    | BEGIN
    | BINARY
    | CLASS
    | CHDIR
    | CHDRIVE
    | COLLECTION
    | DATABASE
    | DELETESETTING
    | ERROR
    | FILECOPY
    | KILL
    | LOAD
    | LIB
    | MID
    | MKDIR
    | NAME
    | OBJECT
    | OUTPUT
    | RANDOM
    | RANDOMIZE
    | READ
    | RESET
    | RMDIR
    | SAVEPICTURE
    | SAVESETTING
    | SENDKEYS
    | SETATTR
    | STEP
    | TEXT
    | TIME
    | UNLOAD
    | VERSION
    | WIDTH
    ;

remKeyword
    : REM
    ;

remStatement
    : REMCOMMENT
    ;

comment
    : COMMENT
    ;

endOfLine
    : WS? (NEWLINE | comment | remStatement) WS?
    ;

endOfStatement
    : (endOfLine | WS? COLON WS?)*
    ;

wsc
    : (WS | LINE_CONTINUATION)+
    ;

/*
* Visual Basic 7.1 Grammar for ANTLR4
*
* Derived from the Visual Basic 7.1 language reference
* https://msopenspecs.azureedge.net/files/MS-VBAL/%5bMS-VBAL%5d.pdf
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

//---------------------------------------------------------------------------------------
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

//---------------------------------------------------------------------------------------
// 5.1 Module Body Structure
// Everything from here down is user generated code.
proceduralModuleBody: proceduralModuleDeclarationSection? proceduralModuleCode;
classModuleBody: classModuleDeclarationSection? classModuleCode;
unrestrictedName
    : reservedIdentifier
    | ambiguousIdentifier
    ;
name
    : untypedName
    | typedName
    ;
untypedName
    : ambiguousIdentifier
    ;

//---------------------------------------------------------------------------------------
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
udtDeclaration: TYPE wsc untypedName endOfStatement+ udtMemberList endOfStatement+ END wsc TYPE;
udtMemberList: udtElement wsc (endOfStatement udtElement)*;
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

// 5.2.4 Class Module Declarations
// 5.2.4.2 Implements Directive
implementsDirective: IMPLEMENTS WS ambiguousIdentifier;

// 5.2.4.3 Event Declaration
eventDeclaration: PUBLIC? wsc EVENT ambiguousIdentifier eventParameterList?;
eventParameterList: '(' wsc? positionalParameters? wsc? ')';


//---------------------------------------------------------------------------------------
// 5.3 Module Code Section Structure
proceduralModuleCode: proceduralModuleCodeElement (endOfLine+ proceduralModuleCodeElement)* endOfLine*;
classModuleCode: classModuleCodeElement (endOfLine+ classModuleCodeElement)* endOfLine*;
proceduralModuleCodeElement: commonModuleCodeElement;
classModuleCodeElement
    : commonModuleCodeElement
    | implementsDirective
    ;
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

// 5.3.1 Procedure Declarations
 subroutine-declaration = procedure-scope [initial-static] 
                "sub" subroutine-name [procedure-parameters] [trailing-static] EOS 
                         [procedure-body EOS] 
                [end-label] "end" "sub" procedure-tail 
  
 function-declaration = procedure-scope [initial-static] 
                         "function" function-name [procedure-parameters] [function-type] [trailing-static] EOS 
                  [procedure-body EOS] 
                   [end-label]  "end" "function" procedure-tail 
  
 property-get-declaration = procedure-scope [initial-static] 
                  "Property" "Get" 
                  function-name [procedure-parameters] [function-type] [trailing-static] EOS 
                            [procedure-body EOS] 
                            [end-label] "end" "property" procedure-tail 
  
propertyLhsDeclaration
    : procedure-scope initialStatic? PROPERTY wsc (LET | SET) subroutineName propertyParameters trailingStatic? endOfStatement
        (procedureBody endOfStatement)?
        endLabel? END wsc PROPERTY procedureTail 
endLabel: statementLabelDefinition;
procedure-Tail
    : wsc NEWLINE
    | commentBody
    | remStatement
    ;

// 5.3.1.1 Procedure Scope
procedureScope
    : PRIVATE
    | PUBLIC
    | FRIEND
    | GLOBAL
    ;

// 5.3.1.2 Static Procedures
initialStatic: STATIC 
trailingStatic: STATIC

// 5.3.1.3 Procedure Names
subroutine-name
    : ambiguousIdentifier
    | prefixedName
    ;
functionName
    : typedName
    | subroutine-name 
    ;
prefixedName
    : eventHandlerName
    | implementedName
    | lifecycleHandlerName
    ;

// 5.3.1.4 Function Type Declarations
functionType = AS wsc typeExpression wsc? arrayDesignator?
arrayDesignator: '(' wsc? ')';

// 5.3.1.5 Parameter Lists

// 5.3.1.8 Event Handler Declarations
eventHandlerName: ambiguousIdentifier;

// 5.3.1.9 Implemented Name Declarations
implementedName: ambiguousIdentifier;

// 5.3.1.10 Lifecycle Handler Declarations
lifecycleHandlerName:
    : CLASS_INITIALIZE
    | CLASS_TERMINATE
    ;

//---------------------------------------------------------------------------------------
// 5.4 Procedure Bodies and Statements
procedureBody: statementBlock;

// 5.4.1 Statement Blocks
statementBlock
    : (blockStatement endOfStatement)*
    ;
blockStatement
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
    
// 5.4.1.1  Statement Labels
statementLabelDefinition: (identifierStatementLabel | lineNumberLabel) ':';
statementLabel
    : identifierStatementLabel
    | lineNumberLabel
    ; 
statementLabelList: statementLabel (wsc? ',' wsc? statementLabel)? 
identifierStatementLabel: ambiguousIdentifier;
lineNumberLabel: (INTEGERLITERAL | SHORTLITERAL);

// 5.4.1.2 Rem Statement
// We have a token for this
remStatement: REMCOMMENT;

// 5.4.2 Control Statements
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

// 5.4.2.1 Call Statement
callStatement
    : CALL WS (simpleNameExpression
        | memberAccessExpression
        | indexExpression
        | withExpression)
    | (simpleNameExpression
        | memberAccessExpression
        | withExpression) argumentList
    ;

// 5.4.2.2 While Statement

// 5.4.2.3 For Statement

// 5.4.2.4 For Each Statement

// 5.4.2.5 Exit For Statement

// 5.4.2.6 Do Statement

// 5.4.2.7 Exit Do Statement

// 5.4.2.8 If Statement

// 5.4.2.9 Single-line If Statement

// 5.4.2.10 Select Case Statement

// 5.4.2.11 Stop Statement

// 5.4.2.12 GoTo Statement

// 5.4.2.13 On…GoTo Statement

// 5.4.2.14 GoSub Statement

// 5.4.2.15 Return Statement
returnStatement: RETURN;

// 5.4.2.16 On…GoSub Statement

// 5.4.2.17 Exit Sub Statement

// 5.4.2.18 Exit Function Statement

// 5.4.2.19 Exit Property Statement

// 5.4.2.20 RaiseEvent Statement

// 5.4.2.21 With Statement

// 5.4.3 Data Manipulation Statements
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

// 5.4.3.1 Local Variable Declarations
localVariableDeclaration: DIM wsc? SHARED? wsc? variableDeclarationList;
staticVariableDeclaration: STATIC wsc variableDeclarationList;

// 5.4.3.2 Local Constant Declarations

// 5.4.3.3 ReDim Statement

// 5.4.3.4 Erase Statement

// 5.4.3.5 Mid/MidB/Mid$/MidB$ Statement

// 5.4.3.6 LSet Statement

// 5.4.3.7 RSet Statement

// 5.4.3.8 Let Statement

// 5.4.3.9 Set Statement

// 5.4.4 Error Handling Statements
errorHandlingStatement
    : onErrorStatement
    | resumeStatement
    | errorStatement
    ;

// 5.4.4.1 On Error Statement

// 5.4.4.2 Resume Statement

// 5.4.4.3 Error Statement

// 5.4.5 File Statements
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

// 5.4.5.1 Open Statement
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

// 5.4.5.1.1 File Numbers
fileNumber
    : markedFileNumber
    | unmarkedFileNumber
    ;
markedFileNumber: '#' expression;
unmarkedFileNumber: expression;

// 5.4.5.2 Close and Reset Statements
closeStatement
    : RESET
    | CLOSE wsc? fileNumberList?
    ;
fileNumberList: fileNumber (wsc? ',' wsc? fileNumber)*;

// 5.4.5.3 Seek Statement
seekStatement: SEEK wsc fileNumber wsc? ',' wsc? position;
position: expression;

// 5.4.5.4 Lock Statement
lockStatement: LOCK wsc fileNumber (wsc? ',' wsc? recordRange);
recordRange
    : startRecordNumber
    | startRecordNumber? wsc TO wsc endRecordNumber
    ;
startRecordNumber: expression;
endRecordNumber: expression;

// 5.4.5.5 Unlock Statement
unlockStatement: UNLOCK wsc fileNumber (wsc? ',' wsc? recordRange)?;

// 5.4.5.6 Line Input Statement
lineInputStatement: LINE wsc INPUT wsc markedFileNumber wsc? ',' wsc? variableName;
variableName: variableExpression;

// 5.4.5.7 Width Statement
widthStatement: WIDTH wsc markedFileNumber wsc? ',' wsc? lineWidth;
lineWidth: expression;

// 5.4.5.8 Print Statement
printStatement: PRINT wsc markedFileNumber wsc? ',' wsc? outputList?;

// 5.4.5.8.1 Output Lists
outputList: outputItem*;
outputItem: outputClause? charPosition?;
outputClause: spcClause | tabClause| outputExpression;
charPosition: ';' | ',';
outputExpression: expression;
spcClause: SPC wsc '(' wsc? spcNumber wsc? ')';
spcNumber: expression;
tabClause: TAB wsc '(' wsc? tabNumber wsc? ')';
tabNumbeer: expression;

// 5.4.5.9 Write Statement
writeStatement: WRITE wsc markedFileNumber wsc? ',' wsc? outputList?;

// 5.4.5.10 Input Statement
inputStatement: INPUT wsc markedFileNumber wsc? ',' wsc? inputList;
inputList: inputVariable (wsc? ',' wsc? inputVariable)*;
inputVariable: boundVariableExpression;

// 5.4.5.11 Put Statement
putStatement: PUT wsc fileNumber wsc? ',' wsc?recordNumber? wsc? ',' data;
recordNumber: expression;
data: expression;

// 5.4.5.12 Get Statement
getStatement: GET wsc fileNumber wsc? ',' wsc? recordNumber? wsc? ',' ws? variable;
variable: variableExpression;

//---------------------------------------------------------------------------------------
// 5.6  Expressions
// Modifying the order will affect the order of operations
// valueExpression must be rolled up into expression due to mutual left recursion
// operatorExpression must be rolled up into expression due to mutual left recursion
expression
    : literalExpression
    | parenthesizedExpression
    | typeOfIsExpression
    | newExpress
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
    
// 5.6.5 Literal Expressions
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

// 5.6.6 Parenthesized Expressions
parenthesizedExpression: LPAREN wsc? expression wsc? RPAREN;

// 5.6.7 TypeOf…Is Expressions
typeofIsExpression = TYPEOF wsc? expression wsc? IS wsc? typeExpression

// 5.6.8 New Expressions
// The name 'newExpression' fails under the Go language
newExpress
    : NEW wsc? expression
    ;

// 5.6.9.8.1 Not Operator
notOperatorExpression: NOT wsc? expression;

// 5.6.9.3.1 Unary - Operator
unaryMinusExpression: MINUS wsc? expression;

// 5.6.10 Simple Name Expressions
simpleNameExpression: name;

// 5.6.11 Instance Expressions
instanceExpression: ME;

// 5.6.12  Member Access Expressions
// This expression is also rolled into lExpression
// changes here must be duplicated there
memberAccessExpression
    : lExpression '.' WS? unrestrictedName
    | lExpression WS? LINE_CONTINUATION WS?'.' WS? unrestrictedName
    ;

// 5.6.13 Index Expressions
indexExpression
    : lExpression WS? '(' WS? argumentList WS ')'
    ;

// 5.6.13.1 Argument Lists
argumentList: positionalOrNamedArgumentList?;
positionalOrNamedArgumentList
    : (positionalArgument WS? ',')* requiredPositionalArgument
    | (positionalArgument WS? ',')* namedArgumentList
    ;
positionalArgument: argumentExpression?;
requiredPositionalArgument: argumentExpression;
namedArgumentList: namedArgument (wsc? ',' wsc? namedArgument)*;
namedArgument: unrestrictedName wsc? ASSIGN wsc? argumentExpression;
argumentExpression
    : BYVAL? wsc? expression
    | addressofExpression
    ;

// 5.6.14 Dictionary Access Expressions

// 5.6.15 With Expressions

// 5.6.16 Constrained Expressions
// The following Expressions have complex static requirements

// 5.6.16.1 Constant Expressions
constantExpression: expression;

// 5.6.16.2 Conditional Compilation Expressions
ccExpression: expression;

// 5.6.16.3 Boolean Expressions
booleanExpression: expression;

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

//---------------------------------------------------------------------------------------
// Many of the following are labeled as tokens in the standard, but are parser rules here.
// 3.3.1 Separator and Special Tokens
// In theory whitespace should be ignored, but there a handful of cases
// where statements MUST be at the beginning of a line or where a NO-WS
// rule appears in the parser rule.
// If may make things simpler her to send all wsc to the hidden channel
// and let a linting tool highlight the couple cases where whitespace
// will cause an error.
wsc: (WS | LINE_CONTINUATION)+;
// known as EOL in MS-VBAL
endOfLine
    : WS? (NEWLINE | commentBody | remStatement)
    ;
// known as EOS in MS-VBAL
endOfStatement
    : (endOfLine | WS? COLON WS?)*
    ;
// The COMMENT token includes the leading single quote
commentBody: COMMENT;

// 3.3.5.2 Reserved Identifiers and IDENTIFIER
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
// Known as IDENTIFIER in MS-VBAL
ambiguousIdentifier
    : IDENTIFIER
    | ambiguousKeyword
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
remKeyword: REM;
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
specialForm
    : ARRAY
    | CIRCLE
    | INPUT
    | INPUTB
    | LBOUND
    | SCALE
    | UBOUND
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
literalIdentifier
    : booleanLiteralIdentifier
    | objectLiteralIdentifier
    | variantLiteralIdentifier
    ;
booleanLiteralIdentifier
    : TRUE
    | FALSE
    ;
objectLiteralIdentifier
    : NOTHING
    ;
variantLiteralIdentifier
    : EMPTY
    | NULL_
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
futureReserved
    : CDECL
    | DECIMAL
    | DEFDEC
    ;

// 3.3.5.3  Special Identifier Forms

// Known as FOREIGN-NAME in MS-VBAL
foreignName: '[' foreignIdentifier ']';
foreignIdentifier: ~(NEWLINE | LINE_CONTINUATION)

// known as BUILTIN-TYPE in MS-VBAL
builtinType
    : reservedTypeIdentifier
    | '[' reservedTypeIdentifier ']'
    | OBJECT
    | '[' OBJECT ']'
    ;

// Known as TYPED-NAME in MS-VBAL
// This probably could be turned into a token
typedName
    : ambiguousIdentifier typeSuffix
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

//---------------------------------------------------------------------------------------
// Extra Rules

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

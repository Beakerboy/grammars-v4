grammar vba_cc;

options {
    caseInsensitive = true;
}

startRule
    : (proceduralModuleHeader | classFileHeader) conditionalModuleBody NEWLINE* EOF
    ;

proceduralModuleHeader
    : 'ATTRIBUTE VB_NAME = ' STRINGLITERAL
    ;

classFileHeader
    : 'VERSION' FLOATLITERAL 'CLASS'?
    ;

// 3.4 Conditional Compilation
conditionalModuleBody: ccBlock+;
ccBlock: (ccConst | ccIfBlock | LOGICAL_LINE)+;

// 3.4.1 Conditional Compilation Const Directive
ccConst: CONST ccVarLhs '=' ccExpression COMMENT?;
ccVarLhs: IDENTIFIER;

// 3.4.2 Conditional Compilation If Directives
ccIfBlock
    : ccIf ccBlock ccElseifBlock* ccElseBlock? ccEndif;
ccIf: IF ccExpression THEN COMMENT?;
ccElseifBlock: ccElseif ccBlock?;
ccElseif: ELSEIF ccExpression THEN COMMENT?;
ccElseBlock: ccElse ccBlock?;
ccElse: ELSE COMMENT?;
ccEndif: ENDIF COMMENT?;
ccExpression
    : literalExpression
    | reservedKeywords
    | IDENTIFIER
    | '-' ccExpression
    | '(' ccExpression ')'
    | ccExpression operator ccExpression
    | ccFunc '(' ccExpression ')'
    ;

// 5.6.5 Literal Expressions
// check on hex and oct
// check definition of integer and float
literalExpression
    : HEXLITERAL
    | OCTLITERAL
    | FLOATLITERAL
    | INTEGERLITERAL
    | STRINGLITERAL
    ;

operator
    : '+'
    | '-'
    | '*'
    | '/'
    | '\\'
    | '^'
    | 'MOD'
    | '&'
    | 'AND'
    | 'OR'
    | 'XOR'
    | 'EQV'
    | 'IMP'
    | 'LIKE'
    | '='
    | '<'
    | '>'
    | '<>'
    | '<='
    | '=>'
    ;

ccFunc
    : 'INT'
    | 'FIX'
    | 'ABS'
    | 'SGN'
    | 'LEN'
    | 'LENB'
    | 'CBOOL'
    | 'CBYTE'
    | 'CCUR'
    | 'CDATE'
    | 'CDBL'
    | 'CINT'
    | 'CLNG'
    | 'CLNGLNG'
    | 'CLNGPTR'
    | 'CSNG'
    | 'CSTR'
    | 'CVAR'
    ;

reservedKeywords
    : WIN16
    | WIN32
    | WIN64
    | VBA6
    | VBA7
    | MAC
    ;

CONST
    : NEWLINE WS? '#CONST'
    ;

IF
    : NEWLINE WS? '#IF'
    ;

ELSEIF
    : NEWLINE WS? '#ELSEIF'
    ;

ELSE
    : NEWLINE WS? '#ELSE'
    ;

ENDIF
    : NEWLINE WS? '#END IF'
    | NEWLINE WS? '#ENDIF'
    ;

THEN
    : 'THEN'
    ;

WIN16
    : 'WIN16'
    ;

WIN32
    : 'WIN32'
    ;

WIN64
    : 'WIN64'
    ;

VBA6
    : 'VBA6'
    ;

VBA7
    : 'VBA7'
    ;

MAC
    : 'MAC'
    ;

IDENTIFIER
    : ~[\]()\r\n\t.,'"|!@#$%^&*\-+:=; ]+
    ;

NEWLINE
    : [\r\n\u2028\u2029]+
    ;

SINGLEQUOTE
    : '\''
    ;

STRINGLITERAL
    : '"' (~["\r\n] | '""')* '"'
    ;

OCTLITERAL
    : '&' [O]? [0-7]+
    ;

HEXLITERAL
    : '&H' [0-9A-F]+
    ;

INTEGERLITERAL
    : (DIGIT DIGIT*
    | HEXLITERAL
    | OCTLITERAL) [%&^]?
    ;

FLOATLITERAL
    : FLOATINGPOINTLITERAL [!#@]?
    | DECIMALLITERAL [!#@]
    ;

fragment FLOATINGPOINTLITERAL
    : DECIMALLITERAL [DE] [+-]? DECIMALLITERAL
    | DECIMALLITERAL '.' DECIMALLITERAL? ([DE] [+-]? DECIMALLITERAL)?
    ;

fragment DECIMALLITERAL
    : DIGIT DIGIT*
    ;

fragment DIGIT
    : [0-9]
    ;

COMMENT
    : SINGLEQUOTE ~[\r\n\u2028\u2029]*
    ;
LOGICAL_LINE
    : NEWLINE WS? ~[\r\n\u2028\u2029#] ~[\r\n\u2028\u2029#]*
    ;
WS
    : ([ \t])+ -> skip
    ;

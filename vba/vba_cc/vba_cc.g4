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

// 5.6.16.2 Conditional Compilation Expressions
ccExpression
    : literalExpression
    | reservedKeywords
    | IDENTIFIER
    | unaryMinusExpression
    | parenthesizedExpression
    | ccExpression operator ccExpression
    | indexExpression
    | notOperatorExpression
    ;

indexExpression
    : ccFunc '(' ccExpression ')'
    ;

parenthesizedExpression
    : '(' ccExpression ')'
    ;

unaryMinusExpression
    :  '-' ccExpression
    ;

notOperatorExpression
    : 'NOT' ccExpression
    ;

literalExpression
    : HEXLITERAL
    | OCTLITERAL
    | FLOATLITERAL
    | INTEGERLITERAL
    | STRINGLITERAL
    | DATELITERAL
    | EMPTY
    | NULL_
    | NOTHING
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
    | '<>' | '><'
    | '<=' | '=<'
    | '=>' | '>='
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

EMPTY
    : 'EMPTY'
    ;

NOTHING
    : 'NOTHING'
    ;

NULL_
    : 'NULL'
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

DATELITERAL
    : '#' DATEORTIME '#'
    ;

fragment DATEORTIME
    : DATEVALUE WS+ TIMEVALUE
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
    : WS+
    | WS? [/,-] WS?
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
    | 'JUNE'
    | 'JULY'
    | 'AUGUST'
    | 'SEPTEMBER'
    | 'OCTOBER'
    | 'NOVEMBER'
    | 'DECEMBER'
    ;

// May has intentionally been left out
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

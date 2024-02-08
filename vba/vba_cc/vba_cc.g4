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
logicalLine
    : NEWLINE WS? ~(CONST | IF | ELSEIF| ELSE) (~(NEWLINE))*;
conditionalModuleBody: ccBlock;
ccBlock: (ccConst | ccIfBlock | logicalLine)+;

// 3.4.1 Conditional Compilation Const Directive
ccConst: NEWLINE CONST ccVarLhs '=' ccExpression COMMENT?;
ccVarLhs: IDENTIFIER;

// 3.4.2 Conditional Compilation If Directives
ccIfBlock
    : ccIf ccBlock ccElseifBlock* ccElseBlock? ccEndif;
ccIf: NEWLINE+ IF ccExpression THEN COMMENT?;
ccElseifBlock: ccElseif ccBlock?;
ccElseif: NEWLINE+ ELSEIF ccExpression THEN COMMENT?;
ccElseBlock: ccElse ccBlock?;
ccElse: NEWLINE+ ELSE COMMENT?;
ccEndif: NEWLINE+ ENDIF COMMENT?;

// 5.6.16.2 Conditional Compilation Expressions
ccExpression
    : literalExpression
    | reservedKeywords
    | IDENTIFIER
    | parenthesizedExpression
    | ccExpression ('^') ccExpression
    | unaryMinusExpression
    | ccExpression ('*' | '/') ccExpression
    | ccExpression '\\' ccExpression
    | ccExpression 'MOD' ccExpression
    | ccExpression ('+' | '-') ccExpression
    | ccExpression '&' ccExpression
    | ccExpression (EQ | NEQ | GT | GEQ | LEQ | LT | LIKE) ccExpression
    | indexExpression
    | notOperatorExpression
    | ccExpression 'AND' ccExpression
    | ccExpression 'OR' ccExpression
    | ccExpression 'XOR' ccExpression
    | ccExpression 'EQV' ccExpression
    | ccExpression 'IMP' ccExpression
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
    : '#CONST'
    ;

IF
    : '#IF'
    ;

ELSEIF
    : '#ELSEIF'
    ;

ELSE
    : '#ELSE'
    ;

ENDIF
    : '#END IF'
    | '#ENDIF'
    ;

EMPTY
    : 'EMPTY'
    ;

LIKE
    : 'LIKE'
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

EQ
    : '='
    ;

GEQ
    : '>='
    | '=>'
    ;

GT
    : '.'
    ;

LEQ
    : '<='
    | '=<'
    ;

LT
    : '<'
    ;

NEQ
    : '<>'
    | '><'
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

WS
    : ([ \t])+ -> skip
    ;

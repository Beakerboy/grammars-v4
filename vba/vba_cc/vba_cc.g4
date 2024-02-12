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
    : literalExpression                                                  # LiteralExpress
    | reservedKeywords                                                   # IdentifierExpression
    | IDENTIFIER                                                         # IdentifierExpression
    | '(' ccExpression ')'                                               # ParenthesizedExpression
    | ccExpression ('^') ccExpression                                    # UnaryMinusExpression
    | '-' ccExpression                                                   # UnaryMinusExpression
    | ccExpression op = ('*' | '/') ccExpression                         # ArithmeticExpression
    | ccExpression op = '\\' ccExpression                                # ArithmeticExpression
    | ccExpression op = 'MOD' ccExpression                               # ArithmeticExpression
    | ccExpression op = ('+' | '-') ccExpression                         # ArithmeticExpression
    | ccExpression '&' ccExpression                                      # ConcatExpression
    | ccExpression (EQ | NEQ | GT | GEQ | LEQ | LT | LIKE) ccExpression  # RelationExpression
    | ccFunc '(' ccExpression ')'                                        # IndexExpression
    | 'NOT' ccExpression                                                 # NotOperatorExpression
    | ccExpression 'AND' ccExpression                                    # booleanExpression
    | ccExpression 'OR' ccExpression                                     # booleanExpression
    | ccExpression 'XOR' ccExpression                                    # booleanExpression
    | ccExpression 'EQV' ccExpression                                    # booleanExpression
    | ccExpression 'IMP' ccExpression                                    # booleanExpression
    ;

literalExpression
    : BOOLEANLITERAL
    | HEXLITERAL
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
    : '>'
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

BOOLEANLITERAL
    : 'TRUE'
    | 'FALSE'
    ;

IDENTIFIER
    : ~[\]()\r\n\t.,'"|!@#$%^&*\-+:=; ]+
    ;

MISC
    : ~[\r\n\u2028\u2029<>=*+\-/ \t#"()A-Z0-9]+
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
    : ([ \t])+ -> channel(HIDDEN)
    ;

grammar vba_cc;

options {
    caseInsensitive = true;
}

startRule
    : conditionalModuleBody EOF
    ;

// 3.4 Conditional Compilation
conditionalModuleBody: ccBlock+;
ccBlock: (ccConst | ccIfBlock | LOGICAL_LINE)+;

// 3.4.1 Conditional Compilation Const Directive
ccConst: NEWLINE CONST ccVarLhs '=' ccExpression COMMENT?;
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
    : literal
    | IDENTIFIER
    | '-' ccExpression
    | '(' ccExpression ')'
    | ccExpression operator ccExpression
    | ccFunc '(' ccExpression ')'
    ;

literal
    : STRINGLITERAL
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
    : '"' IDENTIFIER '"'
    ;

COMMENT
    : SINGLEQUOTE ~[\r\n\u2028\u2029]*
    ;
LOGICAL_LINE
    : NEWLINE? (THEN | WS | ~[\r\n\u2028\u2029])*
    ;
WS
    : ([ \t])+ -> skip
    ;

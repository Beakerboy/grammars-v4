grammar vba_cc;

options {
    caseInsensitive = true;
}

startRule
    : conditionalModuleBody EOF
    ;

// 3.4 Conditional Compilation
conditionalModuleBody: ccBlock*;
ccBlock: (ccConst | ccIfBlock)*;

// 3.4.1 Conditional Compilation Const Directive
ccConst: NEWLINE CONST ccVarLhs EQ ccExpression COMMENT?;
ccVarLhs: IDENTIFIER;

// 3.4.2 Conditional Compilation If Directives
ccIfBlock
    : ccIf ccBlock *ccElseifBlock ccElseBlock? ccEndif;
ccIf: NEWLINE IF ccExpression THEN COMMENT?;
ccElseifBlock: ccElseif ccBlock;
ccElseif: NEWLINE ELSEIF ccExpression THEN COMMENT?;
ccElseBlock: ccElse ccBlock;
ccElse: NEWLINE ELSE ccEol;
ccEndif: NEWLINE ENDIF COMMENT?;
ccExpression
    : literal
    | IDENTIFIER
    | '-' ccExpression
    | '(' ccExpression ')'
    | ccExpression operator ccExpression
    | ccFunc '(' ccExpression ')'
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

WS
    : ([ \t])+ -> skip
    ;

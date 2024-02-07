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
ccConst: NEWLINE CONST ccVarLhs EQ ccExpression COMMENT?
ccVarLhs: name;
ccEol: (COMMENT)?

// 3.4.2 Conditional Compilation If Directives
ccIfBlock
    : ccIf ccBlock *ccElseifBlock ccElseBlock? ccEndif;
ccIf: NEWLINE IF ccExpression THEN COMMENT?;
ccElseifBlock: ccElseif ccBlock;
ccElseif: NEWLINE ELSEIF ccExpression THEN ccEol;
ccElseBlock: ccElse ccBlock;
ccElse: NEWLINE ELSE ccEol;
ccEndif: NEWLINE ENDIF COMMENT?

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

COMMENT
    : SINGLEQUOTE ~[\r\n\u2028\u2029]*
    ;

WS
    : ([ \t])+ -> skip
    ;

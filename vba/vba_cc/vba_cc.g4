grammar vba_cc;

options {
    caseInsensitive = true;
}

startRule
    : conditionalModuleBody EOF
    ;

// 3.4 Conditional Compilation
conditionalModuleBody = ccBlock*
ccBlock: (ccConst | ccIfBlock | logical-line)*

// 3.4.1 Conditional Compilation Const Directive
ccConst: endOfLine '#' CONST cc-var-lhs WS? EQ WS? ccExpression ccEol
ccVarLhs: name;
ccEol: (COMMENT)?

// 3.4.2 Conditional Compilation If Directives
ccIfBlock
    : ccIf ccBlock *ccElseifBlock ccElseBlock? ccEndif;
ccIf: endOfLine IF cc-expression THEN ccEol;
ccElseifBlock: ccElseif ccBlock;
ccElseif: LINE-START ELSEIF ccExpression THEN ccEol;
ccElseBlock: ccElse ccBlock;
ccElse: LINE-START ELSE ccEol;
ccEndif: LINE-START ENDIF ccEol

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

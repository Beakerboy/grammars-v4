grammar vba_like;

options {
    caseInsensitive = false;
}

// module ----------------------------------

program
    : patternElement+
    ;

patternElement
    : CHAR
    | set
    | notSet
    | wildcard
    ;

wildcard
    : WILD_CHAR
    | WILD_SEQ
    | WILD_DIGIT
    ;

notSet
    : '[!' charList? ']'
    ;

charList
    : '-'? charListElement+ '-'?
    ;
    
charListElement
    : CHAR
    | charRange
    | specialChar
    ;

set
    : '[' charList? ']
    ;

charRange
    : CHAR '-' CHAR
    ;
    
specialChar
    : '?'
    | '['
    | '*'
    | '#'
    ;

    

// lexer rules --------------------------------------------------------------------------------

CHAR
    : [A-Z]
    | [a-z]
    | [0-9]
    ;

WILD_CHAR
    : '?'
    ;

WILD_SEQ
    : '*'
    ;

WILD_DIGIT
    : '#'
    ;
    

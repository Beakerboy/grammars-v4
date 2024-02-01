lexer grammar vba_lexer;

options {
    caseInsensitive = true;
}

// keywords
ACCESS
    : 'ACCESS'
    ;

ADDRESSOF
    : 'ADDRESSOF'
    ;

ALIAS
    : 'ALIAS'
    ;

AND
    : 'AND'
    ;

ATTRIBUTE
    : 'ATTRIBUTE'
    ;

APPACTIVATE
    : 'APPACTIVATE'
    ;

APPEND
    : 'APPEND'
    ;

AS
    : 'AS'
    ;

BEGIN
    : 'BEGIN'
    ;

BEEP
    : 'BEEP'
    ;

BINARY
    : 'BINARY'
    ;

BOOLEAN
    : 'BOOLEAN'
    ;

BYVAL
    : 'BYVAL'
    ;

BYREF
    : 'BYREF'
    ;

BYTE
    : 'BYTE'
    ;

CALL
    : 'CALL'
    ;

CASE
    : 'CASE'
    ;

CHDIR
    : 'CHDIR'
    ;

CHDRIVE
    : 'CHDRIVE'
    ;

CLASS
    : 'CLASS'
    ;

CLOSE
    : 'CLOSE'
    ;

COLLECTION
    : 'COLLECTION'
    ;

CONST
    : 'CONST'
    ;

DATABASE
    : 'DATABASE'
    ;

DATE
    : 'DATE'
    ;

DECLARE
    : 'DECLARE'
    ;

DEFBOOL
    : 'DEFBOOL'
    ;

DEFBYTE
    : 'DEFBYTE'
    ;

DEFDATE
    : 'DEFDATE'
    ;

DEFDBL
    : 'DEFDBL'
    ;

DEFDEC
    : 'DEFDEC'
    ;

DEFCUR
    : 'DEFCUR'
    ;

DEFINT
    : 'DEFINT'
    ;

DEFLNG
    : 'DEFLNG'
    ;

DEFOBJ
    : 'DEFOBJ'
    ;

DEFSNG
    : 'DEFSNG'
    ;

DEFSTR
    : 'DEFSTR'
    ;

DEFVAR
    : 'DEFVAR'
    ;

DELETESETTING
    : 'DELETESETTING'
    ;

DIM
    : 'DIM'
    ;

DO
    : 'DO'
    ;

DOUBLE
    : 'DOUBLE'
    ;

EACH
    : 'EACH'
    ;

ELSE
    : 'ELSE'
    ;

ELSEIF
    : 'ELSEIF'
    ;

END_ENUM
    : 'END' WS 'ENUM'
    ;

END_PROPERTY
    : 'END' WS 'PROPERTY'
    ;

END_SELECT
    : 'END' WS 'SELECT'
    ;

END_TYPE
    : 'END' WS 'TYPE'
    ;

END_WITH
    : 'END' WS 'WITH'
    ;

END
    : 'END'
    ;

ENUM
    : 'ENUM'
    ;

EQV
    : 'EQV'
    ;

ERASE
    : 'ERASE'
    ;

ERROR
    : 'ERROR'
    ;

EVENT
    : 'EVENT'
    ;

EXIT_DO
    : 'EXIT' WS 'DO'
    ;

EXIT_FOR
    : 'EXIT' WS 'FOR'
    ;

EXIT_FUNCTION
    : 'EXIT' WS 'FUNCTION'
    ;

EXIT_PROPERTY
    : 'EXIT' WS 'PROPERTY'
    ;

EXIT_SUB
    : 'EXIT' WS 'SUB'
    ;

FALSE
    : 'FALSE'
    ;

FILECOPY
    : 'FILECOPY'
    ;

FRIEND
    : 'FRIEND'
    ;

FOR
    : 'FOR'
    ;

FUNCTION
    : 'FUNCTION'
    ;

GET
    : 'GET'
    ;

GLOBAL
    : 'GLOBAL'
    ;

GOSUB
    : 'GOSUB'
    ;

GOTO
    : 'GOTO'
    ;

IF
    : 'IF'
    ;

IMP
    : 'IMP'
    ;

IMPLEMENTS
    : 'IMPLEMENTS'
    ;

IN
    : 'IN'
    ;

INPUT
    : 'INPUT'
    ;

IS
    : 'IS'
    ;

INTEGER
    : 'INTEGER'
    ;

KILL
    : 'KILL'
    ;

LOAD
    : 'LOAD'
    ;

LOCK
    : 'LOCK'
    ;

LONG
    : 'LONG'
    ;

LOOP
    : 'LOOP'
    ;

LEN
    : 'LEN'
    ;

LET
    : 'LET'
    ;

LIB
    : 'LIB'
    ;

LIKE
    : 'LIKE'
    ;

LINE_INPUT
    : 'LINE' WS 'INPUT'
    ;

LOCK_READ
    : 'LOCK' WS 'READ'
    ;

LOCK_WRITE
    : 'LOCK' WS 'WRITE'
    ;

LOCK_READ_WRITE
    : 'LOCK' WS 'READ' WS 'WRITE'
    ;

LSET
    : 'LSET'
    ;

MACRO_CONST
    : '#CONST'
    ;

MACRO_IF
    : '#IF'
    ;

MACRO_ELSEIF
    : '#ELSEIF'
    ;

MACRO_ELSE
    : '#ELSE'
    ;

MACRO_END_IF
    : '#END' WS? 'IF'
    ;

ME
    : 'ME'
    ;

MID
    : 'MID'
    ;

MKDIR
    : 'MKDIR'
    ;

MOD
    : 'MOD'
    ;

NAME
    : 'NAME'
    ;

NEXT
    : 'NEXT'
    ;

NEW
    : 'NEW'
    ;

NOT
    : 'NOT'
    ;

NOTHING
    : 'NOTHING'
    ;

NULL_
    : 'NULL'
    ;

ON
    : 'ON'
    ;

ON_ERROR
    : 'ON' WS 'ERROR'
    ;

ON_LOCAL_ERROR
    : 'ON' WS 'LOCAL' WS 'ERROR'
    ;

OPEN
    : 'OPEN'
    ;

OPTIONAL
    : 'OPTIONAL'
    ;

OPTION_BASE
    : 'OPTION' WS 'BASE'
    ;

OPTION_EXPLICIT
    : 'OPTION' WS 'EXPLICIT'
    ;

OPTION_COMPARE
    : 'OPTION' WS 'COMPARE'
    ;

OPTION_PRIVATE_MODULE
    : 'OPTION' WS 'PRIVATE' WS 'MODULE'
    ;

OR
    : 'OR'
    ;

OUTPUT
    : 'OUTPUT'
    ;

PARAMARRAY
    : 'PARAMARRAY'
    ;

PRESERVE
    : 'PRESERVE'
    ;

PRINT
    : 'PRINT'
    ;

PRIVATE
    : 'PRIVATE'
    ;

PROPERTY_GET
    : 'PROPERTY' WS 'GET'
    ;

PROPERTY_LET
    : 'PROPERTY' WS 'LET'
    ;

PROPERTY_SET
    : 'PROPERTY' WS 'SET'
    ;

PTRSAFE
    : 'PTRSAFE'
    ;

PUBLIC
    : 'PUBLIC'
    ;

PUT
    : 'PUT'
    ;

RANDOM
    : 'RANDOM'
    ;

RANDOMIZE
    : 'RANDOMIZE'
    ;

RAISEEVENT
    : 'RAISEEVENT'
    ;

READ
    : 'READ'
    ;

READ_WRITE
    : 'READ' WS 'WRITE'
    ;

REDIM
    : 'REDIM'
    ;

REM
    : 'REM'
    ;

RESET
    : 'RESET'
    ;

RESUME
    : 'RESUME'
    ;

RETURN
    : 'RETURN'
    ;

RMDIR
    : 'RMDIR'
    ;

RSET
    : 'RSET'
    ;

SAVEPICTURE
    : 'SAVEPICTURE'
    ;

SAVESETTING
    : 'SAVESETTING'
    ;

SEEK
    : 'SEEK'
    ;

SELECT
    : 'SELECT'
    ;

SENDKEYS
    : 'SENDKEYS'
    ;

SET
    : 'SET'
    ;

SETATTR
    : 'SETATTR'
    ;

SHARED
    : 'SHARED'
    ;

SINGLE
    : 'SINGLE'
    ;

SPC
    : 'SPC'
    ;

STATIC
    : 'STATIC'
    ;

STEP
    : 'STEP'
    ;

STOP
    : 'STOP'
    ;

STRING
    : 'STRING'
    ;

SUB
    : 'SUB'
    ;

TAB
    : 'TAB'
    ;

TEXT
    : 'TEXT'
    ;

THEN
    : 'THEN'
    ;

TIME
    : 'TIME'
    ;

TO
    : 'TO'
    ;

TRUE
    : 'TRUE'
    ;

TYPE
    : 'TYPE'
    ;

TYPEOF
    : 'TYPEOF'
    ;

UNLOAD
    : 'UNLOAD'
    ;

UNLOCK
    : 'UNLOCK'
    ;

UNTIL
    : 'UNTIL'
    ;

VARIANT
    : 'VARIANT'
    ;

VERSION
    : 'VERSION'
    ;

WEND
    : 'WEND'
    ;

WHILE
    : 'WHILE'
    ;

WIDTH
    : 'WIDTH'
    ;

WITH
    : 'WITH'
    ;

WITHEVENTS
    : 'WITHEVENTS'
    ;

WRITE
    : 'WRITE'
    ;

XOR
    : 'XOR'
    ;

// symbols
AMPERSAND
    : '&'
    ;

ASSIGN
    : ':='
    ;

DIV
    : '\\'
    | '/'
    ;

EQ
    : '='
    ;

GEQ
    : '>='
    ;

GT
    : '>'
    ;

LEQ
    : '<='
    ;

LPAREN
    : '('
    ;

LT
    : '<'
    ;

MINUS
    : '-'
    ;

MINUS_EQ
    : '-='
    ;

MULT
    : '*'
    ;

NEQ
    : '<>'
    ;

PLUS
    : '+'
    ;

PLUS_EQ
    : '+='
    ;

POW
    : '^'
    ;

RPAREN
    : ')'
    ;

L_SQUARE_BRACKET
    : '['
    ;

R_SQUARE_BRACKET
    : ']'
    ;

// literals
fragment BLOCK
    : HEXDIGIT HEXDIGIT HEXDIGIT HEXDIGIT
    ;

GUID
    : '{' BLOCK BLOCK MINUS BLOCK MINUS BLOCK MINUS BLOCK MINUS BLOCK BLOCK BLOCK '}'
    ;

STRINGLITERAL
    : '"' (~["\r\n] | '""')* '"'
    ;

OCTLITERAL
    : '&O' [0-7]+ '&'?
    ;

HEXLITERAL
    : '&H' [0-9A-F]+ '&'?
    ;

SHORTLITERAL
    : (PLUS | MINUS)? DIGIT+
    ;

INTEGERLITERAL
    : SHORTLITERAL ('E' SHORTLITERAL)?
    ;

DOUBLELITERAL
    : (PLUS | MINUS)? DIGIT* '.' DIGIT+ ('E' SHORTLITERAL)?
    ;

DATELITERAL
    : '#' DATEORTIME '#'
    ;

fragment DATEORTIME
    : DATEVALUE WS? TIMEVALUE
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
    : WS? [/,-]? WS?
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
    | 'JUNE | AUGUST'
    | 'SEPTEMBER'
    | 'OCTOBER'
    | 'NOVEMBER'
    | 'DECEMBER'
    ;

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

// whitespace, line breaks, comments, ...
LINE_CONTINUATION
    : [ \t] UNDERSCORE '\r'? '\n'
    ;

NEWLINE
    : [\r\n\u2028\u2029]+
    ;

REMCOMMENT
    : COLON? REM WS (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*
    ;

COMMENT
    : SINGLEQUOTE (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*
    ;

SINGLEQUOTE
    : '\''
    ;

COLON
    : ':'
    ;

UNDERSCORE
    : '_'
    ;

WS
    : ([ \t])+
    ;

// identifier
IDENTIFIER
    : ~[\]()\r\n\t.,'"|!@#$%^&*\-+:=; ]+
    | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET
    ;

// letters
fragment LETTER
    : [A-Z_\p{L}]
    ;

fragment DIGIT
    : [0-9]
    ;

fragment HEXDIGIT
    : [A-F0-9]
    ;

fragment LETTERORDIGIT
    : [A-Z0-9_\p{L}]
    ;

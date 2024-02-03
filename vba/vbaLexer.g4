// $antlr-format alignTrailingComments true, columnLimit 150, minEmptyLines 1, maxEmptyLinesToKeep 1, reflowComments false, useTab false
// $antlr-format allowShortRulesOnASingleLine false, allowShortBlocksOnASingleLine true, alignSemicolons hanging, alignColons hanging

lexer grammar vbaLexer;

options {
    caseInsensitive = true;
}

// keywords
ABS
    : 'ABS'
    ;

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

ANY
    : 'ANY'
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

ARRAY
    : 'ARRAY'
    ;

AS
    : 'AS'
    ;

BASE
    : 'BASE'
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

CBOOL
    : 'CBOOL'
    ;

CBYTE
    : 'CBYTE'
    ;

CCUR
    : 'CCUR'
    ;

CDATE
    : 'CDATE'
    ;

CDBL
    : 'CDBL'
    ;

CDEC
    : 'CDEC'
    ;

CDECL
    : 'CDECL'
    ;

CHDIR
    : 'CHDIR'
    ;

CHDRIVE
    : 'CHDRIVE'
    ;

CINT
    : 'CINT'
    ;

CIRCLE
    : 'CIRCLE'
    ;

CLASS
    : 'CLASS'
    ;

CLASS_INITIALIZE
    : 'CLASS_INITIALIZE'
    ;

CLASS_TERMINATE
    : 'CLASS_TERMINATE'
    ;

CLNG
    : 'CLNG'
    ;

CLNGLNG
    : 'CLNGLNG'
    ;

CLNGPTR
    : 'CLNGPTR'
    ;

CLOSE
    : 'CLOSE'
    ;

COLLECTION
    : 'COLLECTION'
    ;

COMPARE
    : 'COMPARE'
    ;

CONST
    : 'CONST'
    ;

CSNG
    : 'CSNG'
    ;

CSTR
    : 'CSTR'
    ;

CVAR
    : 'CVAR'
    ;

CVERR
    : 'CVERR'
    ;

CURRENCY
    : 'CURRENCY'
    ;

DATABASE
    : 'DATABASE'
    ;

DATE
    : 'DATE'
    ;

DEBUG
    : 'DEBUG'
    ;

DECLARE
    : 'DECLARE'
    ;

DECIMAL
    : 'DECIMAL'
    ;

DEFBOOL
    : 'DEFBOOL'
    ;

DEFBYTE
    : 'DEFBYTE'
    ;

DEFCUR
    : 'DEFCUR'
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

DEFINT
    : 'DEFINT'
    ;

DEFLNG
    : 'DEFLNG'
    ;

DEFLNGLNG
    : 'DEFLNGLNG'
    ;

DEFLNGPTR
    : 'DEFLNGPTR'
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

DOEVENTS
    : 'DOEVENTS'
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

EMPTY
    : 'EMPTY'
    ;

ENDIF
    : 'ENDIF'
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

EXIT
    : 'EXIT'
    ;

EXPLICIT
    : 'EXPLICIT'
    ;

FALSE
    : 'FALSE'
    ;

FILECOPY
    : 'FILECOPY'
    ;

FIX
    : 'FIX'
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

GO
    : 'GO'
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

INPUTB
    : 'INPUTB'
    ;

INT
    : 'INT'
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

LBOUND
    : 'LBOUND'
    ;

LEN
    : 'LEN'
    ;

LENB
    : 'LENB'
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

LINE
    : 'LINE'
    ;

LINEINPUT
    : 'LINEINPUT'
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

LONGLONG
    : 'LONGLONG'
    ;

LONGPTR
    : 'LONGPTR'
    ;

LOOP
    : 'LOOP'
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

ME
    : 'ME'
    ;

MID
    : 'MID'
    ;

MIDB
    : 'MIDB'
    ;
MID_D
    : 'MID$'
    ;

MIDB_D
    : 'MIDB$'
    ;

MKDIR
    : 'MKDIR'
    ;

MOD
    : 'MOD'
    ;

MODULE
    : 'MODULE'
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
OBJECT
    : 'OBJECT'
    ;
ON
    : 'ON'
    ;

OPEN
    : 'OPEN'
    ;

OPTION
    : 'OPTION'
    ;

OPTIONAL
    : 'OPTIONAL'
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
PROPERTY
    : 'PROPERTY'
    ;

PSET
    : 'PSET'
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

SCALE
    : 'SCALE'
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

SGN
    : 'SGN'
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

UBOUND
    : 'UBOUND'
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

VB_BASE
    : 'VB_BASE'
    ;

VB_CONTROL
    : 'VB_CONTROL'
    ;

VB_CREATABLE
    : 'VB_CREATABLE'
    ;

VB_CUSTOMIZABLE
    : 'VB_CUSTOMIZABLE'
    ;

VB_DESCRIPTION
    : 'VB_DESCRIPTION'
    ;

VB_EXPOSED
    : 'VB_EXPOSED'
    ;

VB_EXT_KEY 
    : 'VB_EXT_KEY '
    ;

VB_GLOBALNAMESPACE
    : 'VB_GLOBALNAMESPACE'
    ;

VB_HELPID
    : 'VB_HELPID'
    ;

VB_INVOKE_FUNC
    : 'VB_INVOKE_FUNC'
    ;

VB_INVOKE_PROPERTY 
    : 'VB_INVOKE_PROPERTY '
    ;

VB_INVOKE_PROPERTYPUT
    : 'VB_INVOKE_PROPERTYPUT'
    ;

VB_INVOKE_PROPERTYPUTREF
    : 'VB_INVOKE_PROPERTYPUTREF'
    ;

VB_MEMBERFLAGS
    : 'VB_MEMBERFLAGS'
    ;

VB_NAME
    : 'VB_NAME'
    ;

VB_PREDECLAREDID
    : 'VB_PREDECLAREDID'
    ;

VB_PROCDATA
    : 'VB_PROCDATA'
    ;

VB_TEMPLATEDERIVED
    : 'VB_TEMPLATEDERIVED'
    ;

VB_USERMEMID
    : 'VB_USERMEMID'
    ;

VB_VARDESCRIPTION
    : 'VB_VARDESCRIPTION'
    ;

VB_VARHELPID
    : 'VB_VARHELPID'
    ;

VB_VARMEMBERFLAGS
    : 'VB_VARMEMBERFLAGS'
    ;

VB_VARPROCDATA 
    : 'VB_VARPROCDATA '
    ;

VB_VARUSERMEMID
    : 'VB_VARUSERMEMID'
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

ASPERAND
    : '@'
    ;

ASSIGN
    : ':='
    ;

COMMA
    : ','
    ;

DIV
    : '\\'
    | '/'
    ;
Dollar
    : '$'
    ;
EQ
    : '='
    ;
EXCLAM
    : '!'
    ;
GEQ
    : '>='
    ;

GT
    : '>'
    ;

HASH
    : '#'
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

PERCENT
    : '%'
    ;

PERIOD
    : '.'
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

SEMICOLON
    : ';'
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

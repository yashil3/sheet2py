grammar ExcelFormula;

formula: expression EOF;

expression
    : IF '(' expression ',' expression ',' expression ')'       # IfExpr
    | SUM '(' range ')'                                         # SumExpr
    | OR '(' expressionList ')'                                 # OrExpr
    | AND '(' expressionList ')'                                # AndExpr
    | COUNTIF '(' range ',' expression ')'                      # CountIfExpr
    | IFERROR '(' expression ',' expression ')'                 # IfErrorExpr
    | ROWS '(' range ')'                                        # RowsExpr
    | FIND '(' expression ',' expression ')'                    # FindExpr
    | COUNT '(' range ')'                                       # CountExpr
    | VLOOKUP '(' expression ',' range ',' expression ',' expression? ')'  # VLookupExpr
    | ROUNDDOWN '(' expression ',' expression ')'               # RoundDownExpr
    | INDEX '(' range ',' expression ',' expression? ')'        # IndexExpr
    | INDIRECT '(' expression ')'                               # IndirectExpr
    | COUNTIFS '(' range ',' expression (',' range ',' expression)* ')'  # CountIfsExpr
    | EOMONTH '(' expression ',' expression ')'                 # EoMonthExpr
    | NOT '(' expression ')'                                    # NotExpr
    | AVERAGE '(' range ')'                                     # AverageExpr
    | SUMIF '(' range ',' expression ')'                        # SumIfExpr
    | CONCAT '(' expressionList ')'                             # ConcatExpr
    | LEN '(' expression ')'                                    # LenExpr
    | ROUND '(' expression ',' expression ')'                   # RoundExpr
    | ISERROR '(' expression ')'                                # IsErrorExpr
    | YEARFRAC '(' expression ',' expression ')'                # YearfracExpr
    | RIGHT '(' expression ',' expression ')'                   # RightExpr
    | cellReference                                             # CellExpr
    | namedRange                                                # NamedRangeExpr
    | NUMBER                                                    # NumberExpr
    | STRING                                                    # StringExpr
    | expression operator=('*'|'/'|'+'|'-'|'>'|'<'|'>='|'<='|'='|'&') expression  # BinaryOpExpr
    | '(' expression ')'                                        # ParenthesizedExpr
    ;

expressionList : expression (',' expression)*;

range: cellReference ':' cellReference;

cellReference: SHEET_NAME? CELL;

// Named range support
namedRange: NAMED_RANGE_IDENTIFIER;

// Lexer rules
IF      : [Ii][Ff];
SUM     : [Ss][Uu][Mm];
OR      : [Oo][Rr];
AND     : [Aa][Nn][Dd];
COUNTIF : [Cc][Oo][Uu][Nn][Tt][Ii][Ff];
IFERROR : [Ii][Ff][Ee][Rr][Rr][Oo][Rr];
ROWS    : [Rr][Oo][Ww][Ss];
FIND    : [Ff][Ii][Nn][Dd];
COUNT   : [Cc][Oo][Uu][Nn][Tt];
VLOOKUP : [Vv][Ll][Oo][Oo][Kk][Uu][Pp];
ROUNDDOWN : [Rr][Oo][Uu][Nn][Dd][Dd][Oo][Ww][Nn];
INDEX   : [Ii][Nn][Dd][Ee][Xx];
INDIRECT : [Ii][Nn][Dd][Ii][Rr][Ee][Cc][Tt];
COUNTIFS : [Cc][Oo][Uu][Nn][Tt][Ii][Ff][Ss];
EOMONTH : [Ee][Oo][Mm][Oo][Nn][Tt][Hh];
NOT     : [Nn][Oo][Tt];
AVERAGE : [Aa][Vv][Ee][Rr][Aa][Gg][Ee];
SUMIF   : [Ss][Uu][Mm][Ii][Ff];
CONCAT  : [Cc][Oo][Nn][Cc][Aa][Tt];
LEN     : [Ll][Ee][Nn];
ROUND   : [Rr][Oo][Uu][Nn][Dd];
ISERROR : [Ii][Ss][Ee][Rr][Rr][Oo][Rr];
YEARFRAC : [Yy][Ee][Aa][Rr][Ff][Rr][Aa][Cc];
RIGHT   : [Rr][Ii][Gg][Hh][Tt];

// Handle quoted sheet names with brackets and special characters
SHEET_NAME: QUOTED_SHEET_NAME '!' | IDENTIFIER '!';
QUOTED_SHEET_NAME: '\'' (~['\r\n])* '\'';

// Handle absolute references with $ symbols
CELL    : '$'? [A-Z]+ '$'? [0-9]+;

// Named range identifier - must come before IDENTIFIER to avoid conflicts
NAMED_RANGE_IDENTIFIER: [A-Za-z_][A-Za-z0-9_.]*;

NUMBER  : [0-9]+('.'[0-9]+)?;
STRING  : '"' (~'"')* '"';
IDENTIFIER: [a-zA-Z_][a-zA-Z0-9_]*;
WS      : [ \t\r\n]+ -> skip;
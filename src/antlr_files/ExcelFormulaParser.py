# Generated from src/antlr_files/ExcelFormula.g4 by ANTLR 4.13.2
# encoding: utf-8
from antlr4 import *
from io import StringIO
import sys
if sys.version_info[1] > 5:
	from typing import TextIO
else:
	from typing.io import TextIO

def serializedATN():
    return [
        4,1,46,222,2,0,7,0,2,1,7,1,2,2,7,2,2,3,7,3,2,4,7,4,2,5,7,5,1,0,1,
        0,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,3,1,81,8,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,3,1,99,8,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,5,1,118,8,1,10,1,12,1,121,9,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,3,1,193,8,1,1,1,1,1,1,1,5,1,198,8,1,10,1,
        12,1,201,9,1,1,2,1,2,1,2,5,2,206,8,2,10,2,12,2,209,9,2,1,3,1,3,1,
        3,1,3,1,4,3,4,216,8,4,1,4,1,4,1,5,1,5,1,5,0,1,2,6,0,2,4,6,8,10,0,
        1,1,0,4,13,249,0,12,1,0,0,0,2,192,1,0,0,0,4,202,1,0,0,0,6,210,1,
        0,0,0,8,215,1,0,0,0,10,219,1,0,0,0,12,13,3,2,1,0,13,14,5,0,0,1,14,
        1,1,0,0,0,15,16,6,1,-1,0,16,17,5,15,0,0,17,18,5,1,0,0,18,19,3,2,
        1,0,19,20,5,2,0,0,20,21,3,2,1,0,21,22,5,2,0,0,22,23,3,2,1,0,23,24,
        5,3,0,0,24,193,1,0,0,0,25,26,5,16,0,0,26,27,5,1,0,0,27,28,3,6,3,
        0,28,29,5,3,0,0,29,193,1,0,0,0,30,31,5,17,0,0,31,32,5,1,0,0,32,33,
        3,4,2,0,33,34,5,3,0,0,34,193,1,0,0,0,35,36,5,18,0,0,36,37,5,1,0,
        0,37,38,3,4,2,0,38,39,5,3,0,0,39,193,1,0,0,0,40,41,5,19,0,0,41,42,
        5,1,0,0,42,43,3,6,3,0,43,44,5,2,0,0,44,45,3,2,1,0,45,46,5,3,0,0,
        46,193,1,0,0,0,47,48,5,20,0,0,48,49,5,1,0,0,49,50,3,2,1,0,50,51,
        5,2,0,0,51,52,3,2,1,0,52,53,5,3,0,0,53,193,1,0,0,0,54,55,5,21,0,
        0,55,56,5,1,0,0,56,57,3,6,3,0,57,58,5,3,0,0,58,193,1,0,0,0,59,60,
        5,22,0,0,60,61,5,1,0,0,61,62,3,2,1,0,62,63,5,2,0,0,63,64,3,2,1,0,
        64,65,5,3,0,0,65,193,1,0,0,0,66,67,5,23,0,0,67,68,5,1,0,0,68,69,
        3,6,3,0,69,70,5,3,0,0,70,193,1,0,0,0,71,72,5,24,0,0,72,73,5,1,0,
        0,73,74,3,2,1,0,74,75,5,2,0,0,75,76,3,6,3,0,76,77,5,2,0,0,77,78,
        3,2,1,0,78,80,5,2,0,0,79,81,3,2,1,0,80,79,1,0,0,0,80,81,1,0,0,0,
        81,82,1,0,0,0,82,83,5,3,0,0,83,193,1,0,0,0,84,85,5,25,0,0,85,86,
        5,1,0,0,86,87,3,2,1,0,87,88,5,2,0,0,88,89,3,2,1,0,89,90,5,3,0,0,
        90,193,1,0,0,0,91,92,5,26,0,0,92,93,5,1,0,0,93,94,3,6,3,0,94,95,
        5,2,0,0,95,96,3,2,1,0,96,98,5,2,0,0,97,99,3,2,1,0,98,97,1,0,0,0,
        98,99,1,0,0,0,99,100,1,0,0,0,100,101,5,3,0,0,101,193,1,0,0,0,102,
        103,5,27,0,0,103,104,5,1,0,0,104,105,3,2,1,0,105,106,5,3,0,0,106,
        193,1,0,0,0,107,108,5,28,0,0,108,109,5,1,0,0,109,110,3,6,3,0,110,
        111,5,2,0,0,111,119,3,2,1,0,112,113,5,2,0,0,113,114,3,6,3,0,114,
        115,5,2,0,0,115,116,3,2,1,0,116,118,1,0,0,0,117,112,1,0,0,0,118,
        121,1,0,0,0,119,117,1,0,0,0,119,120,1,0,0,0,120,122,1,0,0,0,121,
        119,1,0,0,0,122,123,5,3,0,0,123,193,1,0,0,0,124,125,5,29,0,0,125,
        126,5,1,0,0,126,127,3,2,1,0,127,128,5,2,0,0,128,129,3,2,1,0,129,
        130,5,3,0,0,130,193,1,0,0,0,131,132,5,30,0,0,132,133,5,1,0,0,133,
        134,3,2,1,0,134,135,5,3,0,0,135,193,1,0,0,0,136,137,5,31,0,0,137,
        138,5,1,0,0,138,139,3,6,3,0,139,140,5,3,0,0,140,193,1,0,0,0,141,
        142,5,32,0,0,142,143,5,1,0,0,143,144,3,6,3,0,144,145,5,2,0,0,145,
        146,3,2,1,0,146,147,5,3,0,0,147,193,1,0,0,0,148,149,5,33,0,0,149,
        150,5,1,0,0,150,151,3,4,2,0,151,152,5,3,0,0,152,193,1,0,0,0,153,
        154,5,34,0,0,154,155,5,1,0,0,155,156,3,2,1,0,156,157,5,3,0,0,157,
        193,1,0,0,0,158,159,5,35,0,0,159,160,5,1,0,0,160,161,3,2,1,0,161,
        162,5,2,0,0,162,163,3,2,1,0,163,164,5,3,0,0,164,193,1,0,0,0,165,
        166,5,36,0,0,166,167,5,1,0,0,167,168,3,2,1,0,168,169,5,3,0,0,169,
        193,1,0,0,0,170,171,5,37,0,0,171,172,5,1,0,0,172,173,3,2,1,0,173,
        174,5,2,0,0,174,175,3,2,1,0,175,176,5,3,0,0,176,193,1,0,0,0,177,
        178,5,38,0,0,178,179,5,1,0,0,179,180,3,2,1,0,180,181,5,2,0,0,181,
        182,3,2,1,0,182,183,5,3,0,0,183,193,1,0,0,0,184,193,3,8,4,0,185,
        193,3,10,5,0,186,193,5,43,0,0,187,193,5,44,0,0,188,189,5,1,0,0,189,
        190,3,2,1,0,190,191,5,3,0,0,191,193,1,0,0,0,192,15,1,0,0,0,192,25,
        1,0,0,0,192,30,1,0,0,0,192,35,1,0,0,0,192,40,1,0,0,0,192,47,1,0,
        0,0,192,54,1,0,0,0,192,59,1,0,0,0,192,66,1,0,0,0,192,71,1,0,0,0,
        192,84,1,0,0,0,192,91,1,0,0,0,192,102,1,0,0,0,192,107,1,0,0,0,192,
        124,1,0,0,0,192,131,1,0,0,0,192,136,1,0,0,0,192,141,1,0,0,0,192,
        148,1,0,0,0,192,153,1,0,0,0,192,158,1,0,0,0,192,165,1,0,0,0,192,
        170,1,0,0,0,192,177,1,0,0,0,192,184,1,0,0,0,192,185,1,0,0,0,192,
        186,1,0,0,0,192,187,1,0,0,0,192,188,1,0,0,0,193,199,1,0,0,0,194,
        195,10,2,0,0,195,196,7,0,0,0,196,198,3,2,1,3,197,194,1,0,0,0,198,
        201,1,0,0,0,199,197,1,0,0,0,199,200,1,0,0,0,200,3,1,0,0,0,201,199,
        1,0,0,0,202,207,3,2,1,0,203,204,5,2,0,0,204,206,3,2,1,0,205,203,
        1,0,0,0,206,209,1,0,0,0,207,205,1,0,0,0,207,208,1,0,0,0,208,5,1,
        0,0,0,209,207,1,0,0,0,210,211,3,8,4,0,211,212,5,14,0,0,212,213,3,
        8,4,0,213,7,1,0,0,0,214,216,5,39,0,0,215,214,1,0,0,0,215,216,1,0,
        0,0,216,217,1,0,0,0,217,218,5,41,0,0,218,9,1,0,0,0,219,220,5,42,
        0,0,220,11,1,0,0,0,7,80,98,119,192,199,207,215
    ]

class ExcelFormulaParser ( Parser ):

    grammarFileName = "ExcelFormula.g4"

    atn = ATNDeserializer().deserialize(serializedATN())

    decisionsToDFA = [ DFA(ds, i) for i, ds in enumerate(atn.decisionToState) ]

    sharedContextCache = PredictionContextCache()

    literalNames = [ "<INVALID>", "'('", "','", "')'", "'*'", "'/'", "'+'", 
                     "'-'", "'>'", "'<'", "'>='", "'<='", "'='", "'&'", 
                     "':'" ]

    symbolicNames = [ "<INVALID>", "<INVALID>", "<INVALID>", "<INVALID>", 
                      "<INVALID>", "<INVALID>", "<INVALID>", "<INVALID>", 
                      "<INVALID>", "<INVALID>", "<INVALID>", "<INVALID>", 
                      "<INVALID>", "<INVALID>", "<INVALID>", "IF", "SUM", 
                      "OR", "AND", "COUNTIF", "IFERROR", "ROWS", "FIND", 
                      "COUNT", "VLOOKUP", "ROUNDDOWN", "INDEX", "INDIRECT", 
                      "COUNTIFS", "EOMONTH", "NOT", "AVERAGE", "SUMIF", 
                      "CONCAT", "LEN", "ROUND", "ISERROR", "YEARFRAC", "RIGHT", 
                      "SHEET_NAME", "QUOTED_SHEET_NAME", "CELL", "NAMED_RANGE_IDENTIFIER", 
                      "NUMBER", "STRING", "IDENTIFIER", "WS" ]

    RULE_formula = 0
    RULE_expression = 1
    RULE_expressionList = 2
    RULE_range = 3
    RULE_cellReference = 4
    RULE_namedRange = 5

    ruleNames =  [ "formula", "expression", "expressionList", "range", "cellReference", 
                   "namedRange" ]

    EOF = Token.EOF
    T__0=1
    T__1=2
    T__2=3
    T__3=4
    T__4=5
    T__5=6
    T__6=7
    T__7=8
    T__8=9
    T__9=10
    T__10=11
    T__11=12
    T__12=13
    T__13=14
    IF=15
    SUM=16
    OR=17
    AND=18
    COUNTIF=19
    IFERROR=20
    ROWS=21
    FIND=22
    COUNT=23
    VLOOKUP=24
    ROUNDDOWN=25
    INDEX=26
    INDIRECT=27
    COUNTIFS=28
    EOMONTH=29
    NOT=30
    AVERAGE=31
    SUMIF=32
    CONCAT=33
    LEN=34
    ROUND=35
    ISERROR=36
    YEARFRAC=37
    RIGHT=38
    SHEET_NAME=39
    QUOTED_SHEET_NAME=40
    CELL=41
    NAMED_RANGE_IDENTIFIER=42
    NUMBER=43
    STRING=44
    IDENTIFIER=45
    WS=46

    def __init__(self, input:TokenStream, output:TextIO = sys.stdout):
        super().__init__(input, output)
        self.checkVersion("4.13.2")
        self._interp = ParserATNSimulator(self, self.atn, self.decisionsToDFA, self.sharedContextCache)
        self._predicates = None




    class FormulaContext(ParserRuleContext):
        __slots__ = 'parser'

        def __init__(self, parser, parent:ParserRuleContext=None, invokingState:int=-1):
            super().__init__(parent, invokingState)
            self.parser = parser

        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def EOF(self):
            return self.getToken(ExcelFormulaParser.EOF, 0)

        def getRuleIndex(self):
            return ExcelFormulaParser.RULE_formula

        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterFormula" ):
                listener.enterFormula(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitFormula" ):
                listener.exitFormula(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitFormula" ):
                return visitor.visitFormula(self)
            else:
                return visitor.visitChildren(self)




    def formula(self):

        localctx = ExcelFormulaParser.FormulaContext(self, self._ctx, self.state)
        self.enterRule(localctx, 0, self.RULE_formula)
        try:
            self.enterOuterAlt(localctx, 1)
            self.state = 12
            self.expression(0)
            self.state = 13
            self.match(ExcelFormulaParser.EOF)
        except RecognitionException as re:
            localctx.exception = re
            self._errHandler.reportError(self, re)
            self._errHandler.recover(self, re)
        finally:
            self.exitRule()
        return localctx


    class ExpressionContext(ParserRuleContext):
        __slots__ = 'parser'

        def __init__(self, parser, parent:ParserRuleContext=None, invokingState:int=-1):
            super().__init__(parent, invokingState)
            self.parser = parser


        def getRuleIndex(self):
            return ExcelFormulaParser.RULE_expression

     
        def copyFrom(self, ctx:ParserRuleContext):
            super().copyFrom(ctx)


    class AndExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def AND(self):
            return self.getToken(ExcelFormulaParser.AND, 0)
        def expressionList(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionListContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterAndExpr" ):
                listener.enterAndExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitAndExpr" ):
                listener.exitAndExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitAndExpr" ):
                return visitor.visitAndExpr(self)
            else:
                return visitor.visitChildren(self)


    class StringExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def STRING(self):
            return self.getToken(ExcelFormulaParser.STRING, 0)

        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterStringExpr" ):
                listener.enterStringExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitStringExpr" ):
                listener.exitStringExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitStringExpr" ):
                return visitor.visitStringExpr(self)
            else:
                return visitor.visitChildren(self)


    class IfExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def IF(self):
            return self.getToken(ExcelFormulaParser.IF, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterIfExpr" ):
                listener.enterIfExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitIfExpr" ):
                listener.exitIfExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitIfExpr" ):
                return visitor.visitIfExpr(self)
            else:
                return visitor.visitChildren(self)


    class YearfracExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def YEARFRAC(self):
            return self.getToken(ExcelFormulaParser.YEARFRAC, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterYearfracExpr" ):
                listener.enterYearfracExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitYearfracExpr" ):
                listener.exitYearfracExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitYearfracExpr" ):
                return visitor.visitYearfracExpr(self)
            else:
                return visitor.visitChildren(self)


    class EoMonthExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def EOMONTH(self):
            return self.getToken(ExcelFormulaParser.EOMONTH, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterEoMonthExpr" ):
                listener.enterEoMonthExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitEoMonthExpr" ):
                listener.exitEoMonthExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitEoMonthExpr" ):
                return visitor.visitEoMonthExpr(self)
            else:
                return visitor.visitChildren(self)


    class SumIfExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def SUMIF(self):
            return self.getToken(ExcelFormulaParser.SUMIF, 0)
        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)

        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterSumIfExpr" ):
                listener.enterSumIfExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitSumIfExpr" ):
                listener.exitSumIfExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitSumIfExpr" ):
                return visitor.visitSumIfExpr(self)
            else:
                return visitor.visitChildren(self)


    class IndexExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def INDEX(self):
            return self.getToken(ExcelFormulaParser.INDEX, 0)
        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)

        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterIndexExpr" ):
                listener.enterIndexExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitIndexExpr" ):
                listener.exitIndexExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitIndexExpr" ):
                return visitor.visitIndexExpr(self)
            else:
                return visitor.visitChildren(self)


    class CountIfsExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def COUNTIFS(self):
            return self.getToken(ExcelFormulaParser.COUNTIFS, 0)
        def range_(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.RangeContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,i)

        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterCountIfsExpr" ):
                listener.enterCountIfsExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitCountIfsExpr" ):
                listener.exitCountIfsExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitCountIfsExpr" ):
                return visitor.visitCountIfsExpr(self)
            else:
                return visitor.visitChildren(self)


    class NumberExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def NUMBER(self):
            return self.getToken(ExcelFormulaParser.NUMBER, 0)

        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterNumberExpr" ):
                listener.enterNumberExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitNumberExpr" ):
                listener.exitNumberExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitNumberExpr" ):
                return visitor.visitNumberExpr(self)
            else:
                return visitor.visitChildren(self)


    class VLookupExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def VLOOKUP(self):
            return self.getToken(ExcelFormulaParser.VLOOKUP, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)

        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterVLookupExpr" ):
                listener.enterVLookupExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitVLookupExpr" ):
                listener.exitVLookupExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitVLookupExpr" ):
                return visitor.visitVLookupExpr(self)
            else:
                return visitor.visitChildren(self)


    class NotExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def NOT(self):
            return self.getToken(ExcelFormulaParser.NOT, 0)
        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterNotExpr" ):
                listener.enterNotExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitNotExpr" ):
                listener.exitNotExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitNotExpr" ):
                return visitor.visitNotExpr(self)
            else:
                return visitor.visitChildren(self)


    class NamedRangeExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def namedRange(self):
            return self.getTypedRuleContext(ExcelFormulaParser.NamedRangeContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterNamedRangeExpr" ):
                listener.enterNamedRangeExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitNamedRangeExpr" ):
                listener.exitNamedRangeExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitNamedRangeExpr" ):
                return visitor.visitNamedRangeExpr(self)
            else:
                return visitor.visitChildren(self)


    class RoundExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def ROUND(self):
            return self.getToken(ExcelFormulaParser.ROUND, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterRoundExpr" ):
                listener.enterRoundExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitRoundExpr" ):
                listener.exitRoundExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitRoundExpr" ):
                return visitor.visitRoundExpr(self)
            else:
                return visitor.visitChildren(self)


    class CountIfExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def COUNTIF(self):
            return self.getToken(ExcelFormulaParser.COUNTIF, 0)
        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)

        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterCountIfExpr" ):
                listener.enterCountIfExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitCountIfExpr" ):
                listener.exitCountIfExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitCountIfExpr" ):
                return visitor.visitCountIfExpr(self)
            else:
                return visitor.visitChildren(self)


    class LenExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def LEN(self):
            return self.getToken(ExcelFormulaParser.LEN, 0)
        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterLenExpr" ):
                listener.enterLenExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitLenExpr" ):
                listener.exitLenExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitLenExpr" ):
                return visitor.visitLenExpr(self)
            else:
                return visitor.visitChildren(self)


    class IfErrorExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def IFERROR(self):
            return self.getToken(ExcelFormulaParser.IFERROR, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterIfErrorExpr" ):
                listener.enterIfErrorExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitIfErrorExpr" ):
                listener.exitIfErrorExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitIfErrorExpr" ):
                return visitor.visitIfErrorExpr(self)
            else:
                return visitor.visitChildren(self)


    class SumExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def SUM(self):
            return self.getToken(ExcelFormulaParser.SUM, 0)
        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterSumExpr" ):
                listener.enterSumExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitSumExpr" ):
                listener.exitSumExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitSumExpr" ):
                return visitor.visitSumExpr(self)
            else:
                return visitor.visitChildren(self)


    class OrExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def OR(self):
            return self.getToken(ExcelFormulaParser.OR, 0)
        def expressionList(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionListContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterOrExpr" ):
                listener.enterOrExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitOrExpr" ):
                listener.exitOrExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitOrExpr" ):
                return visitor.visitOrExpr(self)
            else:
                return visitor.visitChildren(self)


    class ConcatExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def CONCAT(self):
            return self.getToken(ExcelFormulaParser.CONCAT, 0)
        def expressionList(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionListContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterConcatExpr" ):
                listener.enterConcatExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitConcatExpr" ):
                listener.exitConcatExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitConcatExpr" ):
                return visitor.visitConcatExpr(self)
            else:
                return visitor.visitChildren(self)


    class AverageExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def AVERAGE(self):
            return self.getToken(ExcelFormulaParser.AVERAGE, 0)
        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterAverageExpr" ):
                listener.enterAverageExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitAverageExpr" ):
                listener.exitAverageExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitAverageExpr" ):
                return visitor.visitAverageExpr(self)
            else:
                return visitor.visitChildren(self)


    class RowsExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def ROWS(self):
            return self.getToken(ExcelFormulaParser.ROWS, 0)
        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterRowsExpr" ):
                listener.enterRowsExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitRowsExpr" ):
                listener.exitRowsExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitRowsExpr" ):
                return visitor.visitRowsExpr(self)
            else:
                return visitor.visitChildren(self)


    class RightExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def RIGHT(self):
            return self.getToken(ExcelFormulaParser.RIGHT, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterRightExpr" ):
                listener.enterRightExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitRightExpr" ):
                listener.exitRightExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitRightExpr" ):
                return visitor.visitRightExpr(self)
            else:
                return visitor.visitChildren(self)


    class CellExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def cellReference(self):
            return self.getTypedRuleContext(ExcelFormulaParser.CellReferenceContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterCellExpr" ):
                listener.enterCellExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitCellExpr" ):
                listener.exitCellExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitCellExpr" ):
                return visitor.visitCellExpr(self)
            else:
                return visitor.visitChildren(self)


    class FindExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def FIND(self):
            return self.getToken(ExcelFormulaParser.FIND, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterFindExpr" ):
                listener.enterFindExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitFindExpr" ):
                listener.exitFindExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitFindExpr" ):
                return visitor.visitFindExpr(self)
            else:
                return visitor.visitChildren(self)


    class RoundDownExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def ROUNDDOWN(self):
            return self.getToken(ExcelFormulaParser.ROUNDDOWN, 0)
        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterRoundDownExpr" ):
                listener.enterRoundDownExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitRoundDownExpr" ):
                listener.exitRoundDownExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitRoundDownExpr" ):
                return visitor.visitRoundDownExpr(self)
            else:
                return visitor.visitChildren(self)


    class IsErrorExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def ISERROR(self):
            return self.getToken(ExcelFormulaParser.ISERROR, 0)
        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterIsErrorExpr" ):
                listener.enterIsErrorExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitIsErrorExpr" ):
                listener.exitIsErrorExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitIsErrorExpr" ):
                return visitor.visitIsErrorExpr(self)
            else:
                return visitor.visitChildren(self)


    class CountExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def COUNT(self):
            return self.getToken(ExcelFormulaParser.COUNT, 0)
        def range_(self):
            return self.getTypedRuleContext(ExcelFormulaParser.RangeContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterCountExpr" ):
                listener.enterCountExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitCountExpr" ):
                listener.exitCountExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitCountExpr" ):
                return visitor.visitCountExpr(self)
            else:
                return visitor.visitChildren(self)


    class ParenthesizedExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterParenthesizedExpr" ):
                listener.enterParenthesizedExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitParenthesizedExpr" ):
                listener.exitParenthesizedExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitParenthesizedExpr" ):
                return visitor.visitParenthesizedExpr(self)
            else:
                return visitor.visitChildren(self)


    class IndirectExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.copyFrom(ctx)

        def INDIRECT(self):
            return self.getToken(ExcelFormulaParser.INDIRECT, 0)
        def expression(self):
            return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,0)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterIndirectExpr" ):
                listener.enterIndirectExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitIndirectExpr" ):
                listener.exitIndirectExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitIndirectExpr" ):
                return visitor.visitIndirectExpr(self)
            else:
                return visitor.visitChildren(self)


    class BinaryOpExprContext(ExpressionContext):

        def __init__(self, parser, ctx:ParserRuleContext): # actually a ExcelFormulaParser.ExpressionContext
            super().__init__(parser)
            self.operator = None # Token
            self.copyFrom(ctx)

        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterBinaryOpExpr" ):
                listener.enterBinaryOpExpr(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitBinaryOpExpr" ):
                listener.exitBinaryOpExpr(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitBinaryOpExpr" ):
                return visitor.visitBinaryOpExpr(self)
            else:
                return visitor.visitChildren(self)



    def expression(self, _p:int=0):
        _parentctx = self._ctx
        _parentState = self.state
        localctx = ExcelFormulaParser.ExpressionContext(self, self._ctx, _parentState)
        _prevctx = localctx
        _startState = 2
        self.enterRecursionRule(localctx, 2, self.RULE_expression, _p)
        self._la = 0 # Token type
        try:
            self.enterOuterAlt(localctx, 1)
            self.state = 192
            self._errHandler.sync(self)
            token = self._input.LA(1)
            if token in [15]:
                localctx = ExcelFormulaParser.IfExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx

                self.state = 16
                self.match(ExcelFormulaParser.IF)
                self.state = 17
                self.match(ExcelFormulaParser.T__0)
                self.state = 18
                self.expression(0)
                self.state = 19
                self.match(ExcelFormulaParser.T__1)
                self.state = 20
                self.expression(0)
                self.state = 21
                self.match(ExcelFormulaParser.T__1)
                self.state = 22
                self.expression(0)
                self.state = 23
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [16]:
                localctx = ExcelFormulaParser.SumExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 25
                self.match(ExcelFormulaParser.SUM)
                self.state = 26
                self.match(ExcelFormulaParser.T__0)
                self.state = 27
                self.range_()
                self.state = 28
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [17]:
                localctx = ExcelFormulaParser.OrExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 30
                self.match(ExcelFormulaParser.OR)
                self.state = 31
                self.match(ExcelFormulaParser.T__0)
                self.state = 32
                self.expressionList()
                self.state = 33
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [18]:
                localctx = ExcelFormulaParser.AndExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 35
                self.match(ExcelFormulaParser.AND)
                self.state = 36
                self.match(ExcelFormulaParser.T__0)
                self.state = 37
                self.expressionList()
                self.state = 38
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [19]:
                localctx = ExcelFormulaParser.CountIfExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 40
                self.match(ExcelFormulaParser.COUNTIF)
                self.state = 41
                self.match(ExcelFormulaParser.T__0)
                self.state = 42
                self.range_()
                self.state = 43
                self.match(ExcelFormulaParser.T__1)
                self.state = 44
                self.expression(0)
                self.state = 45
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [20]:
                localctx = ExcelFormulaParser.IfErrorExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 47
                self.match(ExcelFormulaParser.IFERROR)
                self.state = 48
                self.match(ExcelFormulaParser.T__0)
                self.state = 49
                self.expression(0)
                self.state = 50
                self.match(ExcelFormulaParser.T__1)
                self.state = 51
                self.expression(0)
                self.state = 52
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [21]:
                localctx = ExcelFormulaParser.RowsExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 54
                self.match(ExcelFormulaParser.ROWS)
                self.state = 55
                self.match(ExcelFormulaParser.T__0)
                self.state = 56
                self.range_()
                self.state = 57
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [22]:
                localctx = ExcelFormulaParser.FindExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 59
                self.match(ExcelFormulaParser.FIND)
                self.state = 60
                self.match(ExcelFormulaParser.T__0)
                self.state = 61
                self.expression(0)
                self.state = 62
                self.match(ExcelFormulaParser.T__1)
                self.state = 63
                self.expression(0)
                self.state = 64
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [23]:
                localctx = ExcelFormulaParser.CountExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 66
                self.match(ExcelFormulaParser.COUNT)
                self.state = 67
                self.match(ExcelFormulaParser.T__0)
                self.state = 68
                self.range_()
                self.state = 69
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [24]:
                localctx = ExcelFormulaParser.VLookupExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 71
                self.match(ExcelFormulaParser.VLOOKUP)
                self.state = 72
                self.match(ExcelFormulaParser.T__0)
                self.state = 73
                self.expression(0)
                self.state = 74
                self.match(ExcelFormulaParser.T__1)
                self.state = 75
                self.range_()
                self.state = 76
                self.match(ExcelFormulaParser.T__1)
                self.state = 77
                self.expression(0)
                self.state = 78
                self.match(ExcelFormulaParser.T__1)
                self.state = 80
                self._errHandler.sync(self)
                _la = self._input.LA(1)
                if (((_la) & ~0x3f) == 0 and ((1 << _la) & 34084860428290) != 0):
                    self.state = 79
                    self.expression(0)


                self.state = 82
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [25]:
                localctx = ExcelFormulaParser.RoundDownExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 84
                self.match(ExcelFormulaParser.ROUNDDOWN)
                self.state = 85
                self.match(ExcelFormulaParser.T__0)
                self.state = 86
                self.expression(0)
                self.state = 87
                self.match(ExcelFormulaParser.T__1)
                self.state = 88
                self.expression(0)
                self.state = 89
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [26]:
                localctx = ExcelFormulaParser.IndexExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 91
                self.match(ExcelFormulaParser.INDEX)
                self.state = 92
                self.match(ExcelFormulaParser.T__0)
                self.state = 93
                self.range_()
                self.state = 94
                self.match(ExcelFormulaParser.T__1)
                self.state = 95
                self.expression(0)
                self.state = 96
                self.match(ExcelFormulaParser.T__1)
                self.state = 98
                self._errHandler.sync(self)
                _la = self._input.LA(1)
                if (((_la) & ~0x3f) == 0 and ((1 << _la) & 34084860428290) != 0):
                    self.state = 97
                    self.expression(0)


                self.state = 100
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [27]:
                localctx = ExcelFormulaParser.IndirectExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 102
                self.match(ExcelFormulaParser.INDIRECT)
                self.state = 103
                self.match(ExcelFormulaParser.T__0)
                self.state = 104
                self.expression(0)
                self.state = 105
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [28]:
                localctx = ExcelFormulaParser.CountIfsExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 107
                self.match(ExcelFormulaParser.COUNTIFS)
                self.state = 108
                self.match(ExcelFormulaParser.T__0)
                self.state = 109
                self.range_()
                self.state = 110
                self.match(ExcelFormulaParser.T__1)
                self.state = 111
                self.expression(0)
                self.state = 119
                self._errHandler.sync(self)
                _la = self._input.LA(1)
                while _la==2:
                    self.state = 112
                    self.match(ExcelFormulaParser.T__1)
                    self.state = 113
                    self.range_()
                    self.state = 114
                    self.match(ExcelFormulaParser.T__1)
                    self.state = 115
                    self.expression(0)
                    self.state = 121
                    self._errHandler.sync(self)
                    _la = self._input.LA(1)

                self.state = 122
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [29]:
                localctx = ExcelFormulaParser.EoMonthExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 124
                self.match(ExcelFormulaParser.EOMONTH)
                self.state = 125
                self.match(ExcelFormulaParser.T__0)
                self.state = 126
                self.expression(0)
                self.state = 127
                self.match(ExcelFormulaParser.T__1)
                self.state = 128
                self.expression(0)
                self.state = 129
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [30]:
                localctx = ExcelFormulaParser.NotExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 131
                self.match(ExcelFormulaParser.NOT)
                self.state = 132
                self.match(ExcelFormulaParser.T__0)
                self.state = 133
                self.expression(0)
                self.state = 134
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [31]:
                localctx = ExcelFormulaParser.AverageExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 136
                self.match(ExcelFormulaParser.AVERAGE)
                self.state = 137
                self.match(ExcelFormulaParser.T__0)
                self.state = 138
                self.range_()
                self.state = 139
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [32]:
                localctx = ExcelFormulaParser.SumIfExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 141
                self.match(ExcelFormulaParser.SUMIF)
                self.state = 142
                self.match(ExcelFormulaParser.T__0)
                self.state = 143
                self.range_()
                self.state = 144
                self.match(ExcelFormulaParser.T__1)
                self.state = 145
                self.expression(0)
                self.state = 146
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [33]:
                localctx = ExcelFormulaParser.ConcatExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 148
                self.match(ExcelFormulaParser.CONCAT)
                self.state = 149
                self.match(ExcelFormulaParser.T__0)
                self.state = 150
                self.expressionList()
                self.state = 151
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [34]:
                localctx = ExcelFormulaParser.LenExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 153
                self.match(ExcelFormulaParser.LEN)
                self.state = 154
                self.match(ExcelFormulaParser.T__0)
                self.state = 155
                self.expression(0)
                self.state = 156
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [35]:
                localctx = ExcelFormulaParser.RoundExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 158
                self.match(ExcelFormulaParser.ROUND)
                self.state = 159
                self.match(ExcelFormulaParser.T__0)
                self.state = 160
                self.expression(0)
                self.state = 161
                self.match(ExcelFormulaParser.T__1)
                self.state = 162
                self.expression(0)
                self.state = 163
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [36]:
                localctx = ExcelFormulaParser.IsErrorExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 165
                self.match(ExcelFormulaParser.ISERROR)
                self.state = 166
                self.match(ExcelFormulaParser.T__0)
                self.state = 167
                self.expression(0)
                self.state = 168
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [37]:
                localctx = ExcelFormulaParser.YearfracExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 170
                self.match(ExcelFormulaParser.YEARFRAC)
                self.state = 171
                self.match(ExcelFormulaParser.T__0)
                self.state = 172
                self.expression(0)
                self.state = 173
                self.match(ExcelFormulaParser.T__1)
                self.state = 174
                self.expression(0)
                self.state = 175
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [38]:
                localctx = ExcelFormulaParser.RightExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 177
                self.match(ExcelFormulaParser.RIGHT)
                self.state = 178
                self.match(ExcelFormulaParser.T__0)
                self.state = 179
                self.expression(0)
                self.state = 180
                self.match(ExcelFormulaParser.T__1)
                self.state = 181
                self.expression(0)
                self.state = 182
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [39, 41]:
                localctx = ExcelFormulaParser.CellExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 184
                self.cellReference()
                pass
            elif token in [42]:
                localctx = ExcelFormulaParser.NamedRangeExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 185
                self.namedRange()
                pass
            elif token in [43]:
                localctx = ExcelFormulaParser.NumberExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 186
                self.match(ExcelFormulaParser.NUMBER)
                pass
            elif token in [44]:
                localctx = ExcelFormulaParser.StringExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 187
                self.match(ExcelFormulaParser.STRING)
                pass
            elif token in [1]:
                localctx = ExcelFormulaParser.ParenthesizedExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 188
                self.match(ExcelFormulaParser.T__0)
                self.state = 189
                self.expression(0)
                self.state = 190
                self.match(ExcelFormulaParser.T__2)
                pass
            else:
                raise NoViableAltException(self)

            self._ctx.stop = self._input.LT(-1)
            self.state = 199
            self._errHandler.sync(self)
            _alt = self._interp.adaptivePredict(self._input,4,self._ctx)
            while _alt!=2 and _alt!=ATN.INVALID_ALT_NUMBER:
                if _alt==1:
                    if self._parseListeners is not None:
                        self.triggerExitRuleEvent()
                    _prevctx = localctx
                    localctx = ExcelFormulaParser.BinaryOpExprContext(self, ExcelFormulaParser.ExpressionContext(self, _parentctx, _parentState))
                    self.pushNewRecursionContext(localctx, _startState, self.RULE_expression)
                    self.state = 194
                    if not self.precpred(self._ctx, 2):
                        from antlr4.error.Errors import FailedPredicateException
                        raise FailedPredicateException(self, "self.precpred(self._ctx, 2)")
                    self.state = 195
                    localctx.operator = self._input.LT(1)
                    _la = self._input.LA(1)
                    if not((((_la) & ~0x3f) == 0 and ((1 << _la) & 16368) != 0)):
                        localctx.operator = self._errHandler.recoverInline(self)
                    else:
                        self._errHandler.reportMatch(self)
                        self.consume()
                    self.state = 196
                    self.expression(3) 
                self.state = 201
                self._errHandler.sync(self)
                _alt = self._interp.adaptivePredict(self._input,4,self._ctx)

        except RecognitionException as re:
            localctx.exception = re
            self._errHandler.reportError(self, re)
            self._errHandler.recover(self, re)
        finally:
            self.unrollRecursionContexts(_parentctx)
        return localctx


    class ExpressionListContext(ParserRuleContext):
        __slots__ = 'parser'

        def __init__(self, parser, parent:ParserRuleContext=None, invokingState:int=-1):
            super().__init__(parent, invokingState)
            self.parser = parser

        def expression(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.ExpressionContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.ExpressionContext,i)


        def getRuleIndex(self):
            return ExcelFormulaParser.RULE_expressionList

        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterExpressionList" ):
                listener.enterExpressionList(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitExpressionList" ):
                listener.exitExpressionList(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitExpressionList" ):
                return visitor.visitExpressionList(self)
            else:
                return visitor.visitChildren(self)




    def expressionList(self):

        localctx = ExcelFormulaParser.ExpressionListContext(self, self._ctx, self.state)
        self.enterRule(localctx, 4, self.RULE_expressionList)
        self._la = 0 # Token type
        try:
            self.enterOuterAlt(localctx, 1)
            self.state = 202
            self.expression(0)
            self.state = 207
            self._errHandler.sync(self)
            _la = self._input.LA(1)
            while _la==2:
                self.state = 203
                self.match(ExcelFormulaParser.T__1)
                self.state = 204
                self.expression(0)
                self.state = 209
                self._errHandler.sync(self)
                _la = self._input.LA(1)

        except RecognitionException as re:
            localctx.exception = re
            self._errHandler.reportError(self, re)
            self._errHandler.recover(self, re)
        finally:
            self.exitRule()
        return localctx


    class RangeContext(ParserRuleContext):
        __slots__ = 'parser'

        def __init__(self, parser, parent:ParserRuleContext=None, invokingState:int=-1):
            super().__init__(parent, invokingState)
            self.parser = parser

        def cellReference(self, i:int=None):
            if i is None:
                return self.getTypedRuleContexts(ExcelFormulaParser.CellReferenceContext)
            else:
                return self.getTypedRuleContext(ExcelFormulaParser.CellReferenceContext,i)


        def getRuleIndex(self):
            return ExcelFormulaParser.RULE_range

        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterRange" ):
                listener.enterRange(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitRange" ):
                listener.exitRange(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitRange" ):
                return visitor.visitRange(self)
            else:
                return visitor.visitChildren(self)




    def range_(self):

        localctx = ExcelFormulaParser.RangeContext(self, self._ctx, self.state)
        self.enterRule(localctx, 6, self.RULE_range)
        try:
            self.enterOuterAlt(localctx, 1)
            self.state = 210
            self.cellReference()
            self.state = 211
            self.match(ExcelFormulaParser.T__13)
            self.state = 212
            self.cellReference()
        except RecognitionException as re:
            localctx.exception = re
            self._errHandler.reportError(self, re)
            self._errHandler.recover(self, re)
        finally:
            self.exitRule()
        return localctx


    class CellReferenceContext(ParserRuleContext):
        __slots__ = 'parser'

        def __init__(self, parser, parent:ParserRuleContext=None, invokingState:int=-1):
            super().__init__(parent, invokingState)
            self.parser = parser

        def CELL(self):
            return self.getToken(ExcelFormulaParser.CELL, 0)

        def SHEET_NAME(self):
            return self.getToken(ExcelFormulaParser.SHEET_NAME, 0)

        def getRuleIndex(self):
            return ExcelFormulaParser.RULE_cellReference

        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterCellReference" ):
                listener.enterCellReference(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitCellReference" ):
                listener.exitCellReference(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitCellReference" ):
                return visitor.visitCellReference(self)
            else:
                return visitor.visitChildren(self)




    def cellReference(self):

        localctx = ExcelFormulaParser.CellReferenceContext(self, self._ctx, self.state)
        self.enterRule(localctx, 8, self.RULE_cellReference)
        self._la = 0 # Token type
        try:
            self.enterOuterAlt(localctx, 1)
            self.state = 215
            self._errHandler.sync(self)
            _la = self._input.LA(1)
            if _la==39:
                self.state = 214
                self.match(ExcelFormulaParser.SHEET_NAME)


            self.state = 217
            self.match(ExcelFormulaParser.CELL)
        except RecognitionException as re:
            localctx.exception = re
            self._errHandler.reportError(self, re)
            self._errHandler.recover(self, re)
        finally:
            self.exitRule()
        return localctx


    class NamedRangeContext(ParserRuleContext):
        __slots__ = 'parser'

        def __init__(self, parser, parent:ParserRuleContext=None, invokingState:int=-1):
            super().__init__(parent, invokingState)
            self.parser = parser

        def NAMED_RANGE_IDENTIFIER(self):
            return self.getToken(ExcelFormulaParser.NAMED_RANGE_IDENTIFIER, 0)

        def getRuleIndex(self):
            return ExcelFormulaParser.RULE_namedRange

        def enterRule(self, listener:ParseTreeListener):
            if hasattr( listener, "enterNamedRange" ):
                listener.enterNamedRange(self)

        def exitRule(self, listener:ParseTreeListener):
            if hasattr( listener, "exitNamedRange" ):
                listener.exitNamedRange(self)

        def accept(self, visitor:ParseTreeVisitor):
            if hasattr( visitor, "visitNamedRange" ):
                return visitor.visitNamedRange(self)
            else:
                return visitor.visitChildren(self)




    def namedRange(self):

        localctx = ExcelFormulaParser.NamedRangeContext(self, self._ctx, self.state)
        self.enterRule(localctx, 10, self.RULE_namedRange)
        try:
            self.enterOuterAlt(localctx, 1)
            self.state = 219
            self.match(ExcelFormulaParser.NAMED_RANGE_IDENTIFIER)
        except RecognitionException as re:
            localctx.exception = re
            self._errHandler.reportError(self, re)
            self._errHandler.recover(self, re)
        finally:
            self.exitRule()
        return localctx



    def sempred(self, localctx:RuleContext, ruleIndex:int, predIndex:int):
        if self._predicates == None:
            self._predicates = dict()
        self._predicates[1] = self.expression_sempred
        pred = self._predicates.get(ruleIndex, None)
        if pred is None:
            raise Exception("No predicate with index:" + str(ruleIndex))
        else:
            return pred(localctx, predIndex)

    def expression_sempred(self, localctx:ExpressionContext, predIndex:int):
            if predIndex == 0:
                return self.precpred(self._ctx, 2)
         





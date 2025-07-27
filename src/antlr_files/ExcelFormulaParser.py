# Generated from ExcelFormula.g4 by ANTLR 4.13.2
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
        4,1,41,203,2,0,7,0,2,1,7,1,2,2,7,2,2,3,7,3,2,4,7,4,1,0,1,0,1,0,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,3,1,79,8,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,3,1,97,8,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,5,1,116,8,1,10,1,12,1,119,9,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
        1,1,1,1,1,1,1,1,1,1,3,1,176,8,1,1,1,1,1,1,1,5,1,181,8,1,10,1,12,
        1,184,9,1,1,2,1,2,1,2,5,2,189,8,2,10,2,12,2,192,9,2,1,3,1,3,1,3,
        1,3,1,4,3,4,199,8,4,1,4,1,4,1,4,0,1,2,5,0,2,4,6,8,0,1,1,0,4,12,228,
        0,10,1,0,0,0,2,175,1,0,0,0,4,185,1,0,0,0,6,193,1,0,0,0,8,198,1,0,
        0,0,10,11,3,2,1,0,11,12,5,0,0,1,12,1,1,0,0,0,13,14,6,1,-1,0,14,15,
        5,14,0,0,15,16,5,1,0,0,16,17,3,2,1,0,17,18,5,2,0,0,18,19,3,2,1,0,
        19,20,5,2,0,0,20,21,3,2,1,0,21,22,5,3,0,0,22,176,1,0,0,0,23,24,5,
        15,0,0,24,25,5,1,0,0,25,26,3,6,3,0,26,27,5,3,0,0,27,176,1,0,0,0,
        28,29,5,16,0,0,29,30,5,1,0,0,30,31,3,4,2,0,31,32,5,3,0,0,32,176,
        1,0,0,0,33,34,5,17,0,0,34,35,5,1,0,0,35,36,3,4,2,0,36,37,5,3,0,0,
        37,176,1,0,0,0,38,39,5,18,0,0,39,40,5,1,0,0,40,41,3,6,3,0,41,42,
        5,2,0,0,42,43,3,2,1,0,43,44,5,3,0,0,44,176,1,0,0,0,45,46,5,19,0,
        0,46,47,5,1,0,0,47,48,3,2,1,0,48,49,5,2,0,0,49,50,3,2,1,0,50,51,
        5,3,0,0,51,176,1,0,0,0,52,53,5,20,0,0,53,54,5,1,0,0,54,55,3,6,3,
        0,55,56,5,3,0,0,56,176,1,0,0,0,57,58,5,21,0,0,58,59,5,1,0,0,59,60,
        3,2,1,0,60,61,5,2,0,0,61,62,3,2,1,0,62,63,5,3,0,0,63,176,1,0,0,0,
        64,65,5,22,0,0,65,66,5,1,0,0,66,67,3,6,3,0,67,68,5,3,0,0,68,176,
        1,0,0,0,69,70,5,23,0,0,70,71,5,1,0,0,71,72,3,2,1,0,72,73,5,2,0,0,
        73,74,3,6,3,0,74,75,5,2,0,0,75,76,3,2,1,0,76,78,5,2,0,0,77,79,3,
        2,1,0,78,77,1,0,0,0,78,79,1,0,0,0,79,80,1,0,0,0,80,81,5,3,0,0,81,
        176,1,0,0,0,82,83,5,24,0,0,83,84,5,1,0,0,84,85,3,2,1,0,85,86,5,2,
        0,0,86,87,3,2,1,0,87,88,5,3,0,0,88,176,1,0,0,0,89,90,5,25,0,0,90,
        91,5,1,0,0,91,92,3,6,3,0,92,93,5,2,0,0,93,94,3,2,1,0,94,96,5,2,0,
        0,95,97,3,2,1,0,96,95,1,0,0,0,96,97,1,0,0,0,97,98,1,0,0,0,98,99,
        5,3,0,0,99,176,1,0,0,0,100,101,5,26,0,0,101,102,5,1,0,0,102,103,
        3,2,1,0,103,104,5,3,0,0,104,176,1,0,0,0,105,106,5,27,0,0,106,107,
        5,1,0,0,107,108,3,6,3,0,108,109,5,2,0,0,109,117,3,2,1,0,110,111,
        5,2,0,0,111,112,3,6,3,0,112,113,5,2,0,0,113,114,3,2,1,0,114,116,
        1,0,0,0,115,110,1,0,0,0,116,119,1,0,0,0,117,115,1,0,0,0,117,118,
        1,0,0,0,118,120,1,0,0,0,119,117,1,0,0,0,120,121,5,3,0,0,121,176,
        1,0,0,0,122,123,5,28,0,0,123,124,5,1,0,0,124,125,3,2,1,0,125,126,
        5,2,0,0,126,127,3,2,1,0,127,128,5,3,0,0,128,176,1,0,0,0,129,130,
        5,29,0,0,130,131,5,1,0,0,131,132,3,2,1,0,132,133,5,3,0,0,133,176,
        1,0,0,0,134,135,5,30,0,0,135,136,5,1,0,0,136,137,3,6,3,0,137,138,
        5,3,0,0,138,176,1,0,0,0,139,140,5,31,0,0,140,141,5,1,0,0,141,142,
        3,6,3,0,142,143,5,2,0,0,143,144,3,2,1,0,144,145,5,3,0,0,145,176,
        1,0,0,0,146,147,5,32,0,0,147,148,5,1,0,0,148,149,3,4,2,0,149,150,
        5,3,0,0,150,176,1,0,0,0,151,152,5,33,0,0,152,153,5,1,0,0,153,154,
        3,2,1,0,154,155,5,3,0,0,155,176,1,0,0,0,156,157,5,34,0,0,157,158,
        5,1,0,0,158,159,3,2,1,0,159,160,5,2,0,0,160,161,3,2,1,0,161,162,
        5,3,0,0,162,176,1,0,0,0,163,164,5,35,0,0,164,165,5,1,0,0,165,166,
        3,2,1,0,166,167,5,3,0,0,167,176,1,0,0,0,168,176,3,8,4,0,169,176,
        5,38,0,0,170,176,5,39,0,0,171,172,5,1,0,0,172,173,3,2,1,0,173,174,
        5,3,0,0,174,176,1,0,0,0,175,13,1,0,0,0,175,23,1,0,0,0,175,28,1,0,
        0,0,175,33,1,0,0,0,175,38,1,0,0,0,175,45,1,0,0,0,175,52,1,0,0,0,
        175,57,1,0,0,0,175,64,1,0,0,0,175,69,1,0,0,0,175,82,1,0,0,0,175,
        89,1,0,0,0,175,100,1,0,0,0,175,105,1,0,0,0,175,122,1,0,0,0,175,129,
        1,0,0,0,175,134,1,0,0,0,175,139,1,0,0,0,175,146,1,0,0,0,175,151,
        1,0,0,0,175,156,1,0,0,0,175,163,1,0,0,0,175,168,1,0,0,0,175,169,
        1,0,0,0,175,170,1,0,0,0,175,171,1,0,0,0,176,182,1,0,0,0,177,178,
        10,2,0,0,178,179,7,0,0,0,179,181,3,2,1,3,180,177,1,0,0,0,181,184,
        1,0,0,0,182,180,1,0,0,0,182,183,1,0,0,0,183,3,1,0,0,0,184,182,1,
        0,0,0,185,190,3,2,1,0,186,187,5,2,0,0,187,189,3,2,1,0,188,186,1,
        0,0,0,189,192,1,0,0,0,190,188,1,0,0,0,190,191,1,0,0,0,191,5,1,0,
        0,0,192,190,1,0,0,0,193,194,3,8,4,0,194,195,5,13,0,0,195,196,3,8,
        4,0,196,7,1,0,0,0,197,199,5,36,0,0,198,197,1,0,0,0,198,199,1,0,0,
        0,199,200,1,0,0,0,200,201,5,37,0,0,201,9,1,0,0,0,7,78,96,117,175,
        182,190,198
    ]

class ExcelFormulaParser ( Parser ):

    grammarFileName = "ExcelFormula.g4"

    atn = ATNDeserializer().deserialize(serializedATN())

    decisionsToDFA = [ DFA(ds, i) for i, ds in enumerate(atn.decisionToState) ]

    sharedContextCache = PredictionContextCache()

    literalNames = [ "<INVALID>", "'('", "','", "')'", "'*'", "'/'", "'+'", 
                     "'-'", "'>'", "'<'", "'>='", "'<='", "'='", "':'" ]

    symbolicNames = [ "<INVALID>", "<INVALID>", "<INVALID>", "<INVALID>", 
                      "<INVALID>", "<INVALID>", "<INVALID>", "<INVALID>", 
                      "<INVALID>", "<INVALID>", "<INVALID>", "<INVALID>", 
                      "<INVALID>", "<INVALID>", "IF", "SUM", "OR", "AND", 
                      "COUNTIF", "IFERROR", "ROWS", "FIND", "COUNT", "VLOOKUP", 
                      "ROUNDDOWN", "INDEX", "INDIRECT", "COUNTIFS", "EOMONTH", 
                      "NOT", "AVERAGE", "SUMIF", "CONCAT", "LEN", "ROUND", 
                      "ISERROR", "SHEET_NAME", "CELL", "NUMBER", "STRING", 
                      "IDENTIFIER", "WS" ]

    RULE_formula = 0
    RULE_expression = 1
    RULE_expressionList = 2
    RULE_range = 3
    RULE_cellReference = 4

    ruleNames =  [ "formula", "expression", "expressionList", "range", "cellReference" ]

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
    IF=14
    SUM=15
    OR=16
    AND=17
    COUNTIF=18
    IFERROR=19
    ROWS=20
    FIND=21
    COUNT=22
    VLOOKUP=23
    ROUNDDOWN=24
    INDEX=25
    INDIRECT=26
    COUNTIFS=27
    EOMONTH=28
    NOT=29
    AVERAGE=30
    SUMIF=31
    CONCAT=32
    LEN=33
    ROUND=34
    ISERROR=35
    SHEET_NAME=36
    CELL=37
    NUMBER=38
    STRING=39
    IDENTIFIER=40
    WS=41

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
            self.state = 10
            self.expression(0)
            self.state = 11
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
            self.state = 175
            self._errHandler.sync(self)
            token = self._input.LA(1)
            if token in [14]:
                localctx = ExcelFormulaParser.IfExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx

                self.state = 14
                self.match(ExcelFormulaParser.IF)
                self.state = 15
                self.match(ExcelFormulaParser.T__0)
                self.state = 16
                self.expression(0)
                self.state = 17
                self.match(ExcelFormulaParser.T__1)
                self.state = 18
                self.expression(0)
                self.state = 19
                self.match(ExcelFormulaParser.T__1)
                self.state = 20
                self.expression(0)
                self.state = 21
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [15]:
                localctx = ExcelFormulaParser.SumExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 23
                self.match(ExcelFormulaParser.SUM)
                self.state = 24
                self.match(ExcelFormulaParser.T__0)
                self.state = 25
                self.range_()
                self.state = 26
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [16]:
                localctx = ExcelFormulaParser.OrExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 28
                self.match(ExcelFormulaParser.OR)
                self.state = 29
                self.match(ExcelFormulaParser.T__0)
                self.state = 30
                self.expressionList()
                self.state = 31
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [17]:
                localctx = ExcelFormulaParser.AndExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 33
                self.match(ExcelFormulaParser.AND)
                self.state = 34
                self.match(ExcelFormulaParser.T__0)
                self.state = 35
                self.expressionList()
                self.state = 36
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [18]:
                localctx = ExcelFormulaParser.CountIfExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 38
                self.match(ExcelFormulaParser.COUNTIF)
                self.state = 39
                self.match(ExcelFormulaParser.T__0)
                self.state = 40
                self.range_()
                self.state = 41
                self.match(ExcelFormulaParser.T__1)
                self.state = 42
                self.expression(0)
                self.state = 43
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [19]:
                localctx = ExcelFormulaParser.IfErrorExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 45
                self.match(ExcelFormulaParser.IFERROR)
                self.state = 46
                self.match(ExcelFormulaParser.T__0)
                self.state = 47
                self.expression(0)
                self.state = 48
                self.match(ExcelFormulaParser.T__1)
                self.state = 49
                self.expression(0)
                self.state = 50
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [20]:
                localctx = ExcelFormulaParser.RowsExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 52
                self.match(ExcelFormulaParser.ROWS)
                self.state = 53
                self.match(ExcelFormulaParser.T__0)
                self.state = 54
                self.range_()
                self.state = 55
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [21]:
                localctx = ExcelFormulaParser.FindExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 57
                self.match(ExcelFormulaParser.FIND)
                self.state = 58
                self.match(ExcelFormulaParser.T__0)
                self.state = 59
                self.expression(0)
                self.state = 60
                self.match(ExcelFormulaParser.T__1)
                self.state = 61
                self.expression(0)
                self.state = 62
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [22]:
                localctx = ExcelFormulaParser.CountExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 64
                self.match(ExcelFormulaParser.COUNT)
                self.state = 65
                self.match(ExcelFormulaParser.T__0)
                self.state = 66
                self.range_()
                self.state = 67
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [23]:
                localctx = ExcelFormulaParser.VLookupExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 69
                self.match(ExcelFormulaParser.VLOOKUP)
                self.state = 70
                self.match(ExcelFormulaParser.T__0)
                self.state = 71
                self.expression(0)
                self.state = 72
                self.match(ExcelFormulaParser.T__1)
                self.state = 73
                self.range_()
                self.state = 74
                self.match(ExcelFormulaParser.T__1)
                self.state = 75
                self.expression(0)
                self.state = 76
                self.match(ExcelFormulaParser.T__1)
                self.state = 78
                self._errHandler.sync(self)
                _la = self._input.LA(1)
                if (((_la) & ~0x3f) == 0 and ((1 << _la) & 1099511611394) != 0):
                    self.state = 77
                    self.expression(0)


                self.state = 80
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [24]:
                localctx = ExcelFormulaParser.RoundDownExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 82
                self.match(ExcelFormulaParser.ROUNDDOWN)
                self.state = 83
                self.match(ExcelFormulaParser.T__0)
                self.state = 84
                self.expression(0)
                self.state = 85
                self.match(ExcelFormulaParser.T__1)
                self.state = 86
                self.expression(0)
                self.state = 87
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [25]:
                localctx = ExcelFormulaParser.IndexExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 89
                self.match(ExcelFormulaParser.INDEX)
                self.state = 90
                self.match(ExcelFormulaParser.T__0)
                self.state = 91
                self.range_()
                self.state = 92
                self.match(ExcelFormulaParser.T__1)
                self.state = 93
                self.expression(0)
                self.state = 94
                self.match(ExcelFormulaParser.T__1)
                self.state = 96
                self._errHandler.sync(self)
                _la = self._input.LA(1)
                if (((_la) & ~0x3f) == 0 and ((1 << _la) & 1099511611394) != 0):
                    self.state = 95
                    self.expression(0)


                self.state = 98
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [26]:
                localctx = ExcelFormulaParser.IndirectExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 100
                self.match(ExcelFormulaParser.INDIRECT)
                self.state = 101
                self.match(ExcelFormulaParser.T__0)
                self.state = 102
                self.expression(0)
                self.state = 103
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [27]:
                localctx = ExcelFormulaParser.CountIfsExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 105
                self.match(ExcelFormulaParser.COUNTIFS)
                self.state = 106
                self.match(ExcelFormulaParser.T__0)
                self.state = 107
                self.range_()
                self.state = 108
                self.match(ExcelFormulaParser.T__1)
                self.state = 109
                self.expression(0)
                self.state = 117
                self._errHandler.sync(self)
                _la = self._input.LA(1)
                while _la==2:
                    self.state = 110
                    self.match(ExcelFormulaParser.T__1)
                    self.state = 111
                    self.range_()
                    self.state = 112
                    self.match(ExcelFormulaParser.T__1)
                    self.state = 113
                    self.expression(0)
                    self.state = 119
                    self._errHandler.sync(self)
                    _la = self._input.LA(1)

                self.state = 120
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [28]:
                localctx = ExcelFormulaParser.EoMonthExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 122
                self.match(ExcelFormulaParser.EOMONTH)
                self.state = 123
                self.match(ExcelFormulaParser.T__0)
                self.state = 124
                self.expression(0)
                self.state = 125
                self.match(ExcelFormulaParser.T__1)
                self.state = 126
                self.expression(0)
                self.state = 127
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [29]:
                localctx = ExcelFormulaParser.NotExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 129
                self.match(ExcelFormulaParser.NOT)
                self.state = 130
                self.match(ExcelFormulaParser.T__0)
                self.state = 131
                self.expression(0)
                self.state = 132
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [30]:
                localctx = ExcelFormulaParser.AverageExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 134
                self.match(ExcelFormulaParser.AVERAGE)
                self.state = 135
                self.match(ExcelFormulaParser.T__0)
                self.state = 136
                self.range_()
                self.state = 137
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [31]:
                localctx = ExcelFormulaParser.SumIfExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 139
                self.match(ExcelFormulaParser.SUMIF)
                self.state = 140
                self.match(ExcelFormulaParser.T__0)
                self.state = 141
                self.range_()
                self.state = 142
                self.match(ExcelFormulaParser.T__1)
                self.state = 143
                self.expression(0)
                self.state = 144
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [32]:
                localctx = ExcelFormulaParser.ConcatExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 146
                self.match(ExcelFormulaParser.CONCAT)
                self.state = 147
                self.match(ExcelFormulaParser.T__0)
                self.state = 148
                self.expressionList()
                self.state = 149
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [33]:
                localctx = ExcelFormulaParser.LenExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 151
                self.match(ExcelFormulaParser.LEN)
                self.state = 152
                self.match(ExcelFormulaParser.T__0)
                self.state = 153
                self.expression(0)
                self.state = 154
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [34]:
                localctx = ExcelFormulaParser.RoundExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 156
                self.match(ExcelFormulaParser.ROUND)
                self.state = 157
                self.match(ExcelFormulaParser.T__0)
                self.state = 158
                self.expression(0)
                self.state = 159
                self.match(ExcelFormulaParser.T__1)
                self.state = 160
                self.expression(0)
                self.state = 161
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [35]:
                localctx = ExcelFormulaParser.IsErrorExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 163
                self.match(ExcelFormulaParser.ISERROR)
                self.state = 164
                self.match(ExcelFormulaParser.T__0)
                self.state = 165
                self.expression(0)
                self.state = 166
                self.match(ExcelFormulaParser.T__2)
                pass
            elif token in [36, 37]:
                localctx = ExcelFormulaParser.CellExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 168
                self.cellReference()
                pass
            elif token in [38]:
                localctx = ExcelFormulaParser.NumberExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 169
                self.match(ExcelFormulaParser.NUMBER)
                pass
            elif token in [39]:
                localctx = ExcelFormulaParser.StringExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 170
                self.match(ExcelFormulaParser.STRING)
                pass
            elif token in [1]:
                localctx = ExcelFormulaParser.ParenthesizedExprContext(self, localctx)
                self._ctx = localctx
                _prevctx = localctx
                self.state = 171
                self.match(ExcelFormulaParser.T__0)
                self.state = 172
                self.expression(0)
                self.state = 173
                self.match(ExcelFormulaParser.T__2)
                pass
            else:
                raise NoViableAltException(self)

            self._ctx.stop = self._input.LT(-1)
            self.state = 182
            self._errHandler.sync(self)
            _alt = self._interp.adaptivePredict(self._input,4,self._ctx)
            while _alt!=2 and _alt!=ATN.INVALID_ALT_NUMBER:
                if _alt==1:
                    if self._parseListeners is not None:
                        self.triggerExitRuleEvent()
                    _prevctx = localctx
                    localctx = ExcelFormulaParser.BinaryOpExprContext(self, ExcelFormulaParser.ExpressionContext(self, _parentctx, _parentState))
                    self.pushNewRecursionContext(localctx, _startState, self.RULE_expression)
                    self.state = 177
                    if not self.precpred(self._ctx, 2):
                        from antlr4.error.Errors import FailedPredicateException
                        raise FailedPredicateException(self, "self.precpred(self._ctx, 2)")
                    self.state = 178
                    localctx.operator = self._input.LT(1)
                    _la = self._input.LA(1)
                    if not((((_la) & ~0x3f) == 0 and ((1 << _la) & 8176) != 0)):
                        localctx.operator = self._errHandler.recoverInline(self)
                    else:
                        self._errHandler.reportMatch(self)
                        self.consume()
                    self.state = 179
                    self.expression(3) 
                self.state = 184
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
            self.state = 185
            self.expression(0)
            self.state = 190
            self._errHandler.sync(self)
            _la = self._input.LA(1)
            while _la==2:
                self.state = 186
                self.match(ExcelFormulaParser.T__1)
                self.state = 187
                self.expression(0)
                self.state = 192
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
            self.state = 193
            self.cellReference()
            self.state = 194
            self.match(ExcelFormulaParser.T__12)
            self.state = 195
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
            self.state = 198
            self._errHandler.sync(self)
            _la = self._input.LA(1)
            if _la==36:
                self.state = 197
                self.match(ExcelFormulaParser.SHEET_NAME)


            self.state = 200
            self.match(ExcelFormulaParser.CELL)
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
         





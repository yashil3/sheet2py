# Generated from ExcelFormula.g4 by ANTLR 4.13.2
from antlr4 import *
if "." in __name__:
    from .ExcelFormulaParser import ExcelFormulaParser
else:
    from ExcelFormulaParser import ExcelFormulaParser

# This class defines a complete listener for a parse tree produced by ExcelFormulaParser.
class ExcelFormulaListener(ParseTreeListener):

    # Enter a parse tree produced by ExcelFormulaParser#formula.
    def enterFormula(self, ctx:ExcelFormulaParser.FormulaContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#formula.
    def exitFormula(self, ctx:ExcelFormulaParser.FormulaContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#AndExpr.
    def enterAndExpr(self, ctx:ExcelFormulaParser.AndExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#AndExpr.
    def exitAndExpr(self, ctx:ExcelFormulaParser.AndExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#StringExpr.
    def enterStringExpr(self, ctx:ExcelFormulaParser.StringExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#StringExpr.
    def exitStringExpr(self, ctx:ExcelFormulaParser.StringExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#IfExpr.
    def enterIfExpr(self, ctx:ExcelFormulaParser.IfExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#IfExpr.
    def exitIfExpr(self, ctx:ExcelFormulaParser.IfExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#EoMonthExpr.
    def enterEoMonthExpr(self, ctx:ExcelFormulaParser.EoMonthExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#EoMonthExpr.
    def exitEoMonthExpr(self, ctx:ExcelFormulaParser.EoMonthExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#SumIfExpr.
    def enterSumIfExpr(self, ctx:ExcelFormulaParser.SumIfExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#SumIfExpr.
    def exitSumIfExpr(self, ctx:ExcelFormulaParser.SumIfExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#IndexExpr.
    def enterIndexExpr(self, ctx:ExcelFormulaParser.IndexExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#IndexExpr.
    def exitIndexExpr(self, ctx:ExcelFormulaParser.IndexExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#CountIfsExpr.
    def enterCountIfsExpr(self, ctx:ExcelFormulaParser.CountIfsExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#CountIfsExpr.
    def exitCountIfsExpr(self, ctx:ExcelFormulaParser.CountIfsExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#NumberExpr.
    def enterNumberExpr(self, ctx:ExcelFormulaParser.NumberExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#NumberExpr.
    def exitNumberExpr(self, ctx:ExcelFormulaParser.NumberExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#VLookupExpr.
    def enterVLookupExpr(self, ctx:ExcelFormulaParser.VLookupExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#VLookupExpr.
    def exitVLookupExpr(self, ctx:ExcelFormulaParser.VLookupExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#NotExpr.
    def enterNotExpr(self, ctx:ExcelFormulaParser.NotExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#NotExpr.
    def exitNotExpr(self, ctx:ExcelFormulaParser.NotExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#RoundExpr.
    def enterRoundExpr(self, ctx:ExcelFormulaParser.RoundExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#RoundExpr.
    def exitRoundExpr(self, ctx:ExcelFormulaParser.RoundExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#CountIfExpr.
    def enterCountIfExpr(self, ctx:ExcelFormulaParser.CountIfExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#CountIfExpr.
    def exitCountIfExpr(self, ctx:ExcelFormulaParser.CountIfExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#LenExpr.
    def enterLenExpr(self, ctx:ExcelFormulaParser.LenExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#LenExpr.
    def exitLenExpr(self, ctx:ExcelFormulaParser.LenExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#IfErrorExpr.
    def enterIfErrorExpr(self, ctx:ExcelFormulaParser.IfErrorExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#IfErrorExpr.
    def exitIfErrorExpr(self, ctx:ExcelFormulaParser.IfErrorExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#SumExpr.
    def enterSumExpr(self, ctx:ExcelFormulaParser.SumExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#SumExpr.
    def exitSumExpr(self, ctx:ExcelFormulaParser.SumExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#OrExpr.
    def enterOrExpr(self, ctx:ExcelFormulaParser.OrExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#OrExpr.
    def exitOrExpr(self, ctx:ExcelFormulaParser.OrExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#ConcatExpr.
    def enterConcatExpr(self, ctx:ExcelFormulaParser.ConcatExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#ConcatExpr.
    def exitConcatExpr(self, ctx:ExcelFormulaParser.ConcatExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#AverageExpr.
    def enterAverageExpr(self, ctx:ExcelFormulaParser.AverageExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#AverageExpr.
    def exitAverageExpr(self, ctx:ExcelFormulaParser.AverageExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#RowsExpr.
    def enterRowsExpr(self, ctx:ExcelFormulaParser.RowsExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#RowsExpr.
    def exitRowsExpr(self, ctx:ExcelFormulaParser.RowsExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#CellExpr.
    def enterCellExpr(self, ctx:ExcelFormulaParser.CellExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#CellExpr.
    def exitCellExpr(self, ctx:ExcelFormulaParser.CellExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#FindExpr.
    def enterFindExpr(self, ctx:ExcelFormulaParser.FindExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#FindExpr.
    def exitFindExpr(self, ctx:ExcelFormulaParser.FindExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#RoundDownExpr.
    def enterRoundDownExpr(self, ctx:ExcelFormulaParser.RoundDownExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#RoundDownExpr.
    def exitRoundDownExpr(self, ctx:ExcelFormulaParser.RoundDownExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#IsErrorExpr.
    def enterIsErrorExpr(self, ctx:ExcelFormulaParser.IsErrorExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#IsErrorExpr.
    def exitIsErrorExpr(self, ctx:ExcelFormulaParser.IsErrorExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#CountExpr.
    def enterCountExpr(self, ctx:ExcelFormulaParser.CountExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#CountExpr.
    def exitCountExpr(self, ctx:ExcelFormulaParser.CountExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#ParenthesizedExpr.
    def enterParenthesizedExpr(self, ctx:ExcelFormulaParser.ParenthesizedExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#ParenthesizedExpr.
    def exitParenthesizedExpr(self, ctx:ExcelFormulaParser.ParenthesizedExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#IndirectExpr.
    def enterIndirectExpr(self, ctx:ExcelFormulaParser.IndirectExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#IndirectExpr.
    def exitIndirectExpr(self, ctx:ExcelFormulaParser.IndirectExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#BinaryOpExpr.
    def enterBinaryOpExpr(self, ctx:ExcelFormulaParser.BinaryOpExprContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#BinaryOpExpr.
    def exitBinaryOpExpr(self, ctx:ExcelFormulaParser.BinaryOpExprContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#expressionList.
    def enterExpressionList(self, ctx:ExcelFormulaParser.ExpressionListContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#expressionList.
    def exitExpressionList(self, ctx:ExcelFormulaParser.ExpressionListContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#range.
    def enterRange(self, ctx:ExcelFormulaParser.RangeContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#range.
    def exitRange(self, ctx:ExcelFormulaParser.RangeContext):
        pass


    # Enter a parse tree produced by ExcelFormulaParser#cellReference.
    def enterCellReference(self, ctx:ExcelFormulaParser.CellReferenceContext):
        pass

    # Exit a parse tree produced by ExcelFormulaParser#cellReference.
    def exitCellReference(self, ctx:ExcelFormulaParser.CellReferenceContext):
        pass

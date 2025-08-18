# Generated from src/antlr_files/ExcelFormula.g4 by ANTLR 4.13.2
from antlr4 import *
if "." in __name__:
    from .ExcelFormulaParser import ExcelFormulaParser
else:
    from ExcelFormulaParser import ExcelFormulaParser

# This class defines a complete generic visitor for a parse tree produced by ExcelFormulaParser.

class ExcelFormulaVisitor(ParseTreeVisitor):

    # Visit a parse tree produced by ExcelFormulaParser#formula.
    def visitFormula(self, ctx:ExcelFormulaParser.FormulaContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#AndExpr.
    def visitAndExpr(self, ctx:ExcelFormulaParser.AndExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#StringExpr.
    def visitStringExpr(self, ctx:ExcelFormulaParser.StringExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#IfExpr.
    def visitIfExpr(self, ctx:ExcelFormulaParser.IfExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#YearfracExpr.
    def visitYearfracExpr(self, ctx:ExcelFormulaParser.YearfracExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#EoMonthExpr.
    def visitEoMonthExpr(self, ctx:ExcelFormulaParser.EoMonthExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#SumIfExpr.
    def visitSumIfExpr(self, ctx:ExcelFormulaParser.SumIfExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#IndexExpr.
    def visitIndexExpr(self, ctx:ExcelFormulaParser.IndexExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#CountIfsExpr.
    def visitCountIfsExpr(self, ctx:ExcelFormulaParser.CountIfsExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#NumberExpr.
    def visitNumberExpr(self, ctx:ExcelFormulaParser.NumberExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#VLookupExpr.
    def visitVLookupExpr(self, ctx:ExcelFormulaParser.VLookupExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#NotExpr.
    def visitNotExpr(self, ctx:ExcelFormulaParser.NotExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#NamedRangeExpr.
    def visitNamedRangeExpr(self, ctx:ExcelFormulaParser.NamedRangeExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#RoundExpr.
    def visitRoundExpr(self, ctx:ExcelFormulaParser.RoundExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#CountIfExpr.
    def visitCountIfExpr(self, ctx:ExcelFormulaParser.CountIfExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#LenExpr.
    def visitLenExpr(self, ctx:ExcelFormulaParser.LenExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#IfErrorExpr.
    def visitIfErrorExpr(self, ctx:ExcelFormulaParser.IfErrorExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#SumExpr.
    def visitSumExpr(self, ctx:ExcelFormulaParser.SumExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#OrExpr.
    def visitOrExpr(self, ctx:ExcelFormulaParser.OrExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#ConcatExpr.
    def visitConcatExpr(self, ctx:ExcelFormulaParser.ConcatExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#AverageExpr.
    def visitAverageExpr(self, ctx:ExcelFormulaParser.AverageExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#RowsExpr.
    def visitRowsExpr(self, ctx:ExcelFormulaParser.RowsExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#RightExpr.
    def visitRightExpr(self, ctx:ExcelFormulaParser.RightExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#CellExpr.
    def visitCellExpr(self, ctx:ExcelFormulaParser.CellExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#FindExpr.
    def visitFindExpr(self, ctx:ExcelFormulaParser.FindExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#RoundDownExpr.
    def visitRoundDownExpr(self, ctx:ExcelFormulaParser.RoundDownExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#IsErrorExpr.
    def visitIsErrorExpr(self, ctx:ExcelFormulaParser.IsErrorExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#CountExpr.
    def visitCountExpr(self, ctx:ExcelFormulaParser.CountExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#ParenthesizedExpr.
    def visitParenthesizedExpr(self, ctx:ExcelFormulaParser.ParenthesizedExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#IndirectExpr.
    def visitIndirectExpr(self, ctx:ExcelFormulaParser.IndirectExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#BinaryOpExpr.
    def visitBinaryOpExpr(self, ctx:ExcelFormulaParser.BinaryOpExprContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#expressionList.
    def visitExpressionList(self, ctx:ExcelFormulaParser.ExpressionListContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#range.
    def visitRange(self, ctx:ExcelFormulaParser.RangeContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#cellReference.
    def visitCellReference(self, ctx:ExcelFormulaParser.CellReferenceContext):
        return self.visitChildren(ctx)


    # Visit a parse tree produced by ExcelFormulaParser#namedRange.
    def visitNamedRange(self, ctx:ExcelFormulaParser.NamedRangeContext):
        return self.visitChildren(ctx)



del ExcelFormulaParser
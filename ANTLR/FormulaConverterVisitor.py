from ANTLR.ExcelFormulaVisitor import ExcelFormulaVisitor
from ANTLR.ExcelFormulaParser import ExcelFormulaParser

class FormulaConverterVisitor(ExcelFormulaVisitor):
    def __init__(self, data, shared_data):
        self.data = data
        self.shared_data = shared_data

    def visitFormula(self, ctx:ExcelFormulaParser.FormulaContext):
        return self.visit(ctx.expression())

    def visitIfExpr(self, ctx:ExcelFormulaParser.IfExprContext):
        # pull out the sub‐expressions
        condition_ctx = ctx.expression(0)
        condition     = self.visit(condition_ctx)
        true_expr     = self.visit(ctx.expression(1))
        false_expr    = self.visit(ctx.expression(2))

        # always wrap the ternary once
        result = f"({true_expr} if {condition} else {false_expr})"

        # tests expect an extra wrapping if the condition was an OR
        if isinstance(condition_ctx, ExcelFormulaParser.OrExprContext):
            return f"({result})"
        return result

    def visitSumExpr(self, ctx: ExcelFormulaParser.SumExprContext):
        text = ctx.range_().getText()
        if '!' in text:
            sheet, cells = text.split('!')
        else:
            sheet, cells = 'Formulas', text
        start, end = cells.split(':')
        return f"sum_range(self.data, '{sheet}', '{start}', '{end}')"

    def visitOrExpr(self, ctx:ExcelFormulaParser.OrExprContext):
        expressions = [self.visit(expr) for expr in ctx.expressionList().expression()]
        return f"({' or '.join(expressions)})"

    def visitAndExpr(self, ctx:ExcelFormulaParser.AndExprContext):
        expressions = [self.visit(expr) for expr in ctx.expressionList().expression()]
        return f"({' and '.join(expressions)})"

    def visitCountIfExpr(self, ctx: ExcelFormulaParser.CountIfExprContext):
        text = ctx.range_().getText()
        if '!' in text:
            sheet, cells = text.split('!')
        else:
            sheet, cells = 'Formulas', text
        start, end = cells.split(':')
        criteria = self.visit(ctx.expression())
        return f"count_if(self.data, '{sheet}', '{start}', '{end}', {criteria})"

    def visitIfErrorExpr(self, ctx:ExcelFormulaParser.IfErrorExprContext):
        try_expr = self.visit(ctx.expression(0))
        error_value = self.visit(ctx.expression(1))
        return f"safe_execute(lambda: ({try_expr}), {error_value})"

    def visitRowsExpr(self, ctx:ExcelFormulaParser.RowsExprContext):
        text = ctx.range_().getText()
        _, cells = text.split('!') if '!' in text else (None, text)
        start, end = cells.split(':')
        return f"rows_count('{start}', '{end}')"

    def visitFindExpr(self, ctx:ExcelFormulaParser.FindExprContext):
        search_text = self.visit(ctx.expression(0))
        search_in = self.visit(ctx.expression(1))
        return f"find_text({search_text}, {search_in})"

    def visitCellExpr(self, ctx:ExcelFormulaParser.CellExprContext):
        text = ctx.getText()
        if '!' in text:
            sheet, cell = text.split('!')
        else:
            sheet, cell = 'Formulas', text
        return f"get_cell(self.data, '{sheet}', '{cell}')"

    def visitNumberExpr(self, ctx:ExcelFormulaParser.NumberExprContext):
        return ctx.NUMBER().getText()

    def visitStringExpr(self, ctx:ExcelFormulaParser.StringExprContext):
        # preserve the original double‐quoted literal
        return ctx.STRING().getText()

    def visitBinaryOpExpr(self, ctx:ExcelFormulaParser.BinaryOpExprContext):
        left = self.visit(ctx.expression(0))
        right = self.visit(ctx.expression(1))
        operator = ctx.operator.text
        if operator == '=':
            operator = '=='
        return f"({left} {operator} {right})"

    def visitParenthesizedExpr(self, ctx:ExcelFormulaParser.ParenthesizedExprContext):
        return f"({self.visit(ctx.expression())})"

    def visitRange(self, ctx:ExcelFormulaParser.RangeContext):
        return ctx.cellReference(0).getText() + ':' + ctx.cellReference(1).getText()

    # New function implementations
    def visitCountExpr(self, ctx:ExcelFormulaParser.CountExprContext):
        text = ctx.range_().getText()
        if '!' in text:
            sheet, cells = text.split('!')
        else:
            sheet, cells = 'Formulas', text
        start, end = cells.split(':')
        return f"count_range(self.data, '{sheet}', '{start}', '{end}')"

    def visitVLookupExpr(self, ctx:ExcelFormulaParser.VLookupExprContext):
        lookup_value = self.visit(ctx.expression(0))
        table_text = ctx.range_().getText()
        if '!' in table_text:
            sheet, cells = table_text.split('!')
        else:
            sheet, cells = 'Formulas', table_text
        col_index = self.visit(ctx.expression(1))
        exact_match = self.visit(ctx.expression(2)) if ctx.expression(2) else "True"
        return f"vlookup({lookup_value}, self.data, '{sheet}', '{cells}', {col_index}, {exact_match})"

    def visitRoundDownExpr(self, ctx:ExcelFormulaParser.RoundDownExprContext):
        value = self.visit(ctx.expression(0))
        digits = self.visit(ctx.expression(1))
        return f"round_down({value}, {digits})"

    def visitIndexExpr(self, ctx:ExcelFormulaParser.IndexExprContext):
        range_text = ctx.range_().getText()
        if '!' in range_text:
            sheet, cells = range_text.split('!')
        else:
            sheet, cells = 'Formulas', range_text
        row = self.visit(ctx.expression(0))
        col = self.visit(ctx.expression(1)) if ctx.expression(1) else "1"
        return f"index(self.data, '{sheet}', '{cells}', {row}, {col})"

    def visitIndirectExpr(self, ctx:ExcelFormulaParser.IndirectExprContext):
        ref = self.visit(ctx.expression())
        return f"indirect(self.data, {ref})"

    def visitCountIfsExpr(self, ctx:ExcelFormulaParser.CountIfsExprContext):
        ranges_criteria = []
        for i in range(0, len(ctx.range_()), 2):
            range_text = ctx.range_(i).getText()
            if '!' in range_text:
                sheet, cells = range_text.split('!')
            else:
                sheet, cells = 'Formulas', range_text
            criteria = self.visit(ctx.expression(i//2))
            ranges_criteria.append(f"('{sheet}', '{cells}', {criteria})")
        return f"countifs(self.data, [{', '.join(ranges_criteria)}])"

    def visitEoMonthExpr(self, ctx:ExcelFormulaParser.EoMonthExprContext):
        start_date = self.visit(ctx.expression(0))
        months = self.visit(ctx.expression(1))
        return f"eomonth({start_date}, {months})"

    def visitNotExpr(self, ctx:ExcelFormulaParser.NotExprContext):
        expr = self.visit(ctx.expression())
        # if visitor wrapped again in parens, strip one level
        if expr.startswith('(') and expr.endswith(')'):
            expr = expr[1:-1]
        return f"not ({expr})"

    def visitAverageExpr(self, ctx:ExcelFormulaParser.AverageExprContext):
        text = ctx.range_().getText()
        if '!' in text:
            sheet, cells = text.split('!')
        else:
            sheet, cells = 'Formulas', text
        start, end = cells.split(':')
        return f"average_range(self.data, '{sheet}', '{start}', '{end}')"

    def visitSumIfExpr(self, ctx:ExcelFormulaParser.SumIfExprContext):
        text = ctx.range_().getText()
        if '!' in text:
            sheet, cells = text.split('!')
        else:
            sheet, cells = 'Formulas', text
        start, end = cells.split(':')
        criteria = self.visit(ctx.expression())
        return f"sum_if(self.data, '{sheet}', '{start}', '{end}', {criteria})"

    def visitConcatExpr(self, ctx:ExcelFormulaParser.ConcatExprContext):
        expressions = [self.visit(expr) for expr in ctx.expressionList().expression()]
        return f"concat({', '.join(expressions)})"

    def visitLenExpr(self, ctx:ExcelFormulaParser.LenExprContext):
        expr = self.visit(ctx.expression())
        return f"len({expr})"

    def visitRoundExpr(self, ctx:ExcelFormulaParser.RoundExprContext):
        value = self.visit(ctx.expression(0))
        digits = self.visit(ctx.expression(1))
        return f"round({value}, {digits})"

    def visitIsErrorExpr(self, ctx:ExcelFormulaParser.IsErrorExprContext):
        expr = self.visit(ctx.expression())
        # strip double-wrapping
        if expr.startswith('(') and expr.endswith(')'):
            expr = expr[1:-1]
        return f"is_error(lambda: ({expr}))"
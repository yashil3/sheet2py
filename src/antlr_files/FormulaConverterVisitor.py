from src.antlr_files.ExcelFormulaVisitor import ExcelFormulaVisitor
from src.antlr_files.ExcelFormulaParser import ExcelFormulaParser


class FormulaConverterVisitor(ExcelFormulaVisitor):
    def __init__(self, data, shared_data, sheet_name):
        self.data = data
        self.shared_data = shared_data
        self.dependencies = set()
        self.sheet_name = sheet_name
        # Track which semantic keys are used by the formula
        self.input_keys = set()

    def _extract_sheet_and_cells(self, text):
        """Helper method to extract sheet name and cells from range text, handling quoted sheet names."""
        if '!' in text:
            sheet, cells = text.split('!')
            # Handle quoted sheet names by removing quotes
            if sheet.startswith("'") and sheet.endswith("'"):
                sheet = sheet[1:-1]
        else:
            sheet, cells = self.sheet_name, text
        return sheet, cells

    def _maybe_key_lookup(self, sheet, cell_ref):
        """If a semantic key exists for this cell, return get_value expression; else None."""
        try:
            cell_to_key_map = self.shared_data.get('cell_to_key_map', {})
            key = cell_to_key_map.get(sheet, {}).get(cell_ref)
            if key:
                self.input_keys.add((sheet, key))
                # Use get_value for semantic access
                return f"get_value(data, '{sheet}', {repr(key)})"
        except Exception:
            pass
        return None

    # --- Helpers for ranges -> keys ---
    def _col_to_num(self, col):
        n = 0
        for c in col:
            if 'A' <= c <= 'Z':
                n = n * 26 + (ord(c) - ord('A') + 1)
            elif 'a' <= c <= 'z':
                n = n * 26 + (ord(c) - ord('a') + 1)
        return n

    def _num_to_col(self, num):
        s = ''
        while num:
            num, rem = divmod(num - 1, 26)
            s = chr(ord('A') + rem) + s
        return s

    def _parse_cell(self, ref):
        letters = ''.join([ch for ch in ref if ch.isalpha()])
        digits = ''.join([ch for ch in ref if ch.isdigit()])
        return letters, int(digits)

    def _expand_range_cells(self, start, end):
        sc, sr = self._parse_cell(start)
        ec, er = self._parse_cell(end)
        scn, ecn = self._col_to_num(sc), self._col_to_num(ec)
        for r in range(sr, er + 1):
            for cn in range(scn, ecn + 1):
                yield f"{self._num_to_col(cn)}{r}"

    def _keys_for_range(self, sheet, start, end):
        cell_to_key_map = self.shared_data.get('cell_to_key_map', {}).get(sheet, {})
        cells = list(self._expand_range_cells(start, end))
        keys = []
        for c in cells:
            k = cell_to_key_map.get(c)
            if k:
                keys.append(k)
            else:
                # If any cell lacks a key, we consider this not a pure key-based range
                return None
        return keys if keys else None

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
        sheet, cells = self._extract_sheet_and_cells(text)
        start, end = cells.split(':')
        self.dependencies.add(f"{sheet}!{start}:{end}")
        keys = self._keys_for_range(sheet, start, end)
        if keys:
            # Track all semantic keys used
            for k in keys:
                self.input_keys.add((sheet, k))
            return f"sum_keys(data, '{sheet}', {repr(keys)})"
        return f"sum_range(data, '{sheet}', '{start}', '{end}')"

    def visitOrExpr(self, ctx:ExcelFormulaParser.OrExprContext):
        expressions = [self.visit(expr) for expr in ctx.expressionList().expression()]
        return f"({' or '.join(expressions)})"

    def visitAndExpr(self, ctx:ExcelFormulaParser.AndExprContext):
        expressions = [self.visit(expr) for expr in ctx.expressionList().expression()]
        return f"({' and '.join(expressions)})"

    def visitCountIfExpr(self, ctx: ExcelFormulaParser.CountIfExprContext):
        text = ctx.range_().getText()
        sheet, cells = self._extract_sheet_and_cells(text)
        start, end = cells.split(':')
        self.dependencies.add(f"{sheet}!{start}:{end}")
        criteria = self.visit(ctx.expression())
        keys = self._keys_for_range(sheet, start, end)
        if keys:
            for k in keys:
                self.input_keys.add((sheet, k))
            return f"count_if_keys(data, '{sheet}', {repr(keys)}, {criteria})"
        return f"count_if(data, '{sheet}', '{start}', '{end}', {criteria})"

    def visitIfErrorExpr(self, ctx:ExcelFormulaParser.IfErrorExprContext):
        try_expr = self.visit(ctx.expression(0))
        error_value = self.visit(ctx.expression(1))
        return f"safe_execute(lambda: ({try_expr}), {error_value})"

    def visitRowsExpr(self, ctx:ExcelFormulaParser.RowsExprContext):
        text = ctx.range_().getText()
        sheet, cells = self._extract_sheet_and_cells(text)
        start, end = cells.split(':')
        self.dependencies.add(f"{sheet}!{start}:{end}")
        return f"rows_count('{start}', '{end}')"

    def visitFindExpr(self, ctx:ExcelFormulaParser.FindExprContext):
        search_text = self.visit(ctx.expression(0))
        search_in = self.visit(ctx.expression(1))
        return f"find_text({search_text}, {search_in})"

    def visitCellExpr(self, ctx:ExcelFormulaParser.CellExprContext):
        text = ctx.getText()
        sheet, cell = self._extract_sheet_and_cells(text)
        self.dependencies.add(f"{sheet}!{cell}")
        # Prefer key-based lookup when available
        key_lookup = self._maybe_key_lookup(sheet, cell)
        if key_lookup:
            return key_lookup
        return f"get_cell(data, '{sheet}', '{cell}')"

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
        elif operator == '&':
            # Convert Excel concatenation to Python string concatenation
            return f"str({left}) + str({right})"
        return f"({left} {operator} {right})"

    def visitParenthesizedExpr(self, ctx:ExcelFormulaParser.ParenthesizedExprContext):
        return f"({self.visit(ctx.expression())})"

    def visitRange(self, ctx:ExcelFormulaParser.RangeContext):
        return ctx.cellReference(0).getText() + ':' + ctx.cellReference(1).getText()

    # New function implementations
    def visitCountExpr(self, ctx:ExcelFormulaParser.CountExprContext):
        text = ctx.range_().getText()
        sheet, cells = self._extract_sheet_and_cells(text)
        start, end = cells.split(':')
        return f"count_range(data, '{sheet}', '{start}', '{end}')"

    def visitVLookupExpr(self, ctx:ExcelFormulaParser.VLookupExprContext):
        lookup_value = self.visit(ctx.expression(0))
        table_text = ctx.range_().getText()
        sheet, cells = self._extract_sheet_and_cells(table_text)
        col_index = self.visit(ctx.expression(1))
        exact_match = self.visit(ctx.expression(2)) if ctx.expression(2) else "True"
        return f"vlookup({lookup_value}, data, '{sheet}', '{cells}', {col_index}, {exact_match})"

    def visitRoundDownExpr(self, ctx:ExcelFormulaParser.RoundDownExprContext):
        value = self.visit(ctx.expression(0))
        digits = self.visit(ctx.expression(1))
        return f"round_down({value}, {digits})"

    def visitIndexExpr(self, ctx:ExcelFormulaParser.IndexExprContext):
        range_text = ctx.range_().getText()
        sheet, cells = self._extract_sheet_and_cells(range_text)
        row = self.visit(ctx.expression(0))
        col = self.visit(ctx.expression(1)) if ctx.expression(1) else "1"
        return f"index(data, '{sheet}', '{cells}', {row}, {col})"

    def visitIndirectExpr(self, ctx:ExcelFormulaParser.IndirectExprContext):
        ref = self.visit(ctx.expression())
        return f"indirect(data, {ref})"

    def visitCountIfsExpr(self, ctx:ExcelFormulaParser.CountIfsExprContext):
        ranges_criteria = []
        for i in range(0, len(ctx.range_()), 2):
            range_text = ctx.range_(i).getText()
            sheet, cells = self._extract_sheet_and_cells(range_text)
            criteria = self.visit(ctx.expression(i//2))
            ranges_criteria.append(f"('{sheet}', '{cells}', {criteria})")
        return f"countifs(data, [{', '.join(ranges_criteria)}])"

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
        sheet, cells = self._extract_sheet_and_cells(text)
        start, end = cells.split(':')
        keys = self._keys_for_range(sheet, start, end)
        if keys:
            for k in keys:
                self.input_keys.add((sheet, k))
            return f"average_keys(data, '{sheet}', {repr(keys)})"
        return f"average_range(data, '{sheet}', '{start}', '{end}')"

    def visitSumIfExpr(self, ctx:ExcelFormulaParser.SumIfExprContext):
        text = ctx.range_().getText()
        sheet, cells = self._extract_sheet_and_cells(text)
        start, end = cells.split(':')
        criteria = self.visit(ctx.expression())
        keys = self._keys_for_range(sheet, start, end)
        if keys:
            for k in keys:
                self.input_keys.add((sheet, k))
            return f"sum_if_keys(data, '{sheet}', {repr(keys)}, {criteria})"
        return f"sum_if(data, '{sheet}', '{start}', '{end}', {criteria})"

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
    
    def visitYearfracExpr(self, ctx:ExcelFormulaParser.YearfracExprContext):
        start_date = self.visit(ctx.expression(0))
        end_date = self.visit(ctx.expression(1))
        return f"yearfrac({start_date}, {end_date})"
    
    def visitRightExpr(self, ctx:ExcelFormulaParser.RightExprContext):
        text = self.visit(ctx.expression(0))
        num_chars = self.visit(ctx.expression(1))
        return f"right_text({text}, {num_chars})"
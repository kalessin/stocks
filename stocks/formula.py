import re


_CELL_RE = re.compile(r"\[\.([A-Z]+)(\d+)\]")
_RANGE_RE = re.compile(r"\[\.([A-Z]+)(\d+):\.([A-Z]+)(\d+)\]")
_REPLACEMENTS = [
    ('of:=', ''),
    ('SUM', 'sum'),
]


def evaluate(formula_string, celldict):
    """
    >>> celldict = {'K3': 4, 'K4': 1, 'K5': -2, 'K7': 8, 'K8': 2, 'K9': -5}
    >>> evaluate('of:=SUM([.K3:.K5])', celldict)
    3
    >>> evaluate('of:=[.K5]-[.K7]', celldict)
    -10
    >>> evaluate('of:=[.K3]+SUM([.K7:.K9])', celldict)
    9
    """
    for m in list(_RANGE_RE.finditer(formula_string))[::-1]:
        colstart, rowstart, colend, rowend = m.groups()
        assert colstart == colend, 'for now not supported multicolumn formula'
        vals = []
        for row in range(int(rowstart), int(rowend) + 1):
            vals.append(celldict[f'{colstart}{row}'])
        formula_string = formula_string[:m.start()] + repr(vals) + formula_string[m.end():]
    for m in list(_CELL_RE.finditer(formula_string))[::-1]:
        col, row = m.groups()
        val = celldict[f'{col}{row}']
        formula_string = formula_string[:m.start()] + repr(val) + formula_string[m.end():]
    for op, repl in _REPLACEMENTS:
        formula_string = formula_string.replace(op, repl)
    try:
        return eval(formula_string)
    except ZeroDivisionError:
        print(f"Error evaluating {formula_string}")
        return "#DIV/0!"
    except TypeError as e:
        print(f"Error evaluating {formula_string}")
        if "#DIV/0!" in formula_string:
            return "#DIV/0!"
        raise
    except:
        print(f"Error evaluating {formula_string}")
        raise

    raise ValueError(f"Formula '{formula_string}' could not be evaluated")

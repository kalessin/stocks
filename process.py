#!/usr/bin/env python3
import re
import json
import argparse
from collections import defaultdict

from odf.opendocument import load
from odf.table import TableRow, TableCell
from odf.text import P

from stocks.formula import evaluate

_FILE_RE = re.compile(r"(\w+)-(\w+)-(\w+)")
_NUM_ROWS = 152

_TRANSLATION = {
    'income_statement': {
        'Revenue': 3,
        'CostOfRevenue': 7,
        'SellingGeneralAdministrative': 10,
        'ResearchDevelopment': 11,
        'InterestExpense': 20,
        'OtherExpenseIncome': 22,
        'IncomeBeforeIncomeTaxes': 23,
        'ConsolidatedNetIncomeLoss': 26,
        'NoncontrollingInterestIncome': 28,
        'PreferredStockDividendsDeclared': 37,
        'NetIncomeLossAttributableCommonShareholders': 38,
        'dilutedsharesoutstanding': 40,
        'CommonStockDividendsDeclared': 41,
    },
    'balance_sheet_statement': {
        'cashequivalents': 46,
        'shortterminvestments': 47,
        'tradereceivables': 50,
        'inventories': 53,
        'prepaidexpensesothercurrentassets': 54,
        'propertyplantequipment': 58,
        'propertyplantequipmentnet': 60,
        'goodwill': 61,
        'intangibleassetsnet': 62,
        'longterminvestments': 63,
        'otherassets': 64,
        'accountspayable': 67,
        'accruedothercurrentliabilities': 68,
        'debtcurrent': 70,
        'longtermdebtcapitalleaseobligations': 76,
        'taxespayablenoncurrent': 79,
        'othernoncurrentliabilities': 81,
        'preferredstock': 85,
        'commonstock': 86,
        'additionalpaidincapital': 87,
        'retainedearningsdeficit': 88,
        'accumulatedothercomprehensiveincome': 90,
    },
    'cash_flow_statement': {
    },
}


_DIVISORS = defaultdict(lambda: 1000000.0)
_DIVISORS.update({
    'CommonStockDividendsDeclared': 1.0,
    'PreferredStockDividendsDeclared': 1.0,
})


def _get_column_ord(col, colord=1):
    """
    >>> _get_column_ord('A')
    1
    >>> _get_column_ord('K')
    11
    >>> _get_column_ord('Z')
    26
    >>> _get_column_ord('AA')
    27
    >>> _get_column_ord('AB')
    28
    >>> _get_column_ord('AZ')
    52
    >>> _get_column_ord('BA')
    53
    >>> _get_column_ord('BZ')
    78
    >>> _get_column_ord('CA')
    79
    """
    if len(col) == 0:
        return colord
    firstord = ord(col[0]) - ord('A')
    if len(col) == 1:
        return colord + firstord
    return _get_column_ord(col[1:], 26 * (firstord + 1) + 1)


def _get_column_from_ord(ordinal):
    """
    >>> samples = ['A', 'Z', 'AA', 'AB', 'AZ', 'BA', 'BZ', 'CA']
    >>> [_get_column_from_ord(_get_column_ord(s)) for s in samples] == samples
    True
    """
    if ordinal <= 26:
        return chr(ordinal + 64)
    return _get_column_from_ord((ordinal - 1) // 26) + _get_column_from_ord((ordinal - 1) % 26 + 1)


def _incr_column(column):
    """
    >>> _incr_column('A')
    'B'
    >>> _incr_column('Z')
    'AA'
    >>> _incr_column('AA')
    'AB'
    >>> _incr_column('AZ')
    'BA'
    """
    return _get_column_from_ord(_get_column_ord(column) + 1)


def _getCell(sheet, rowel, col):
    colord = _get_column_ord(col)
    coln = 0
    for cell in rowel.getElementsByType(TableCell):
        n = int(cell.getAttribute('numbercolumnsrepeated') or 1)
        if coln + n > colord:
            if colord - coln == 1:
                # we are in the target cell, just remove the repeated attribute
                # and insert a new cell with repeated decreased by 1
                cell.removeAttribute('numbercolumnsrepeated')
                newcell = TableCell(stylename=cell.getAttribute('stylename'), numbercolumnsrepeated=n - 1)
                rowel.insertBefore(newcell, cell.nextSibling)
                return cell
            else:
                # we are in a previous cell, just remove the repeated, and add two new cells,
                # one being the target cell and another with repeated decreased by two
                cell.removeAttribute('numbercolumnsrepeated')
                nextSibling = cell.nextSibling
                newcell = TableCell(stylename=cell.getAttribute('stylename'))
                rowel.insertBefore(newcell, nextSibling)
                if n > 2:
                    newnewcell = TableCell(stylename=cell.getAttribute('stylename'))
                    if n > 3:
                        newnewcell.setAttribute('numbercolumnsrepeated', n - 2)
                    rowel.insertBefore(newnewcell, nextSibling)
                return newcell
        elif coln + n == colord:
            return cell
        coln += n
    raise ValueError


def _getRowByIndex(sheet, idx):
    assert idx > 0, f"Wrong row index: {idx}"
    rown = 0
    for row in sheet.getElementsByType(TableRow):
        n = int(row.getAttribute('numberrowsrepeated') or 1)
        if rown + n > idx:
            if idx - rown == 1:
                # we are in the target row, just remove the repeated attribute
                # and insert a new row with repeated decreased by 1
                row.removeAttribute('numberrowsrepeated')
                newrow = TableRow(stylename=row.getAttribute('stylename'), numberrowsrepeated=n - 1)
                sheet.insertBefore(newrow, row.nextSibling)
                return row
            else:
                # we are in a previous row, just remove the repeated, and add two new rows,
                # one being the target row and another with repeated decreased by two
                row.removeAttribute('numberrowsrepeated')
                nextSibling = row.nextSibling
                newrow = TableRow(stylename=row.getAttribute('stylename'))
                sheet.insertBefore(newrow, nextSibling)
                if n > 2:
                    newnewrow = TableRow(stylename=row.getAttribute('stylename'))
                    if n > 3:
                        newnewrow.setAttribute('numberrowsrepeated', n - 2)
                    sheet.insertBefore(newnewrow, nextSibling)
                return newrow
        elif rown + n == idx:
            return row
        rown += n
    raise ValueError(f"Error retrieving row {idx}")


class Process(object):
    def __init__(self):

        parser = argparse.ArgumentParser()
        parser.add_argument('spreadsheet', help='Update given spreadsheet')
        parser.add_argument('ifile', help='Process given input json (generated by spider)')
        parser.add_argument('column', help='Column where to start to add new data.')

        self.args = parser.parse_args()
        self.__filled_cells = {}

    def run(self):
        doc = load(self.args.spreadsheet)
        company, statement, period_type = _FILE_RE.match(self.args.ifile).groups()
        for sheet in doc.spreadsheet.childNodes:
            if sheet.getAttribute('name') == company:
                translations = _TRANSLATION[statement]
                new_data = json.load(open(self.args.ifile))
                column = self.args.column
                for fundamental in new_data['fundamentals'][::-1]:
                    if period_type == 'annual' and not fundamental['filing_type'].startswith('10-K'):
                        continue
                    for tag in fundamental['tags']:
                        if tag['tag'] in translations:
                            value = tag['value'] / _DIVISORS[tag['tag']]
                            if value:
                                row = translations[tag['tag']]
                                rowel = _getRowByIndex(sheet, row)
                                cell = _getCell(sheet, rowel, column)
                                cell.setAttribute("value", value)
                                self.__filled_cells[f'{column}{row}'] = value
                                cell.setAttribute("valuetype", 'float')
                                if cell.getAttribute("formula"):
                                    cell.removeAttribute("formula")
                                if cell.firstChild is not None:
                                    cell.removeChild(cell.firstChild)
                                cell.addElement(P(text=value))
                                print(f"Updated cell {column}{row} with value {value}")

                    for row in range(1, _NUM_ROWS + 1):
                        rowel = _getRowByIndex(sheet, row)
                        try:
                            cell = _getCell(sheet, rowel, column)
                        except ValueError:
                            print(f"Cell {column}{row} is not initialized.")
                        else:
                            formula = cell.getAttribute('formula')
                            if formula is not None:
                                try:
                                    value = evaluate(formula, self.__filled_cells)
                                except Exception as e:
                                    print(f'Error evaluating cell {column}{row}: {formula}')
                                    print(f'{e!r}')
                                    return
                                print(f'Evaluated cell {column}{row} with result {value} ({formula})')
                                self.__filled_cells[f'{column}{row}'] = value
                                cell.setAttribute('value', value)
                                if cell.firstChild is not None:
                                    cell.removeChild(cell.firstChild)

                    column = _incr_column(column)

                doc.save(self.args.spreadsheet)
                print("Saved", self.args.spreadsheet)
                break
        else:
            print(f"No sheet with name {company}")


if __name__ == '__main__':
    process = Process()
    process.run()

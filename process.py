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
_MAX_NUM_ROWS = 200


_TRANSLATION = {
    'income_statement': {
        'revenue': 3,
        'costofrevenue': 7,
        'sellinggeneraladministrative': 10,
        'researchdevelopment': 11,
        'depreciationdepletionamortizationexpense': 12,
        'restructuringimpairmentchargesincomeopex': 13,
        'otheroperatingexpenses': 16,
        'interestexpense': 21,
        'otherexpenseincome': 23,
        'incomebeforeincometaxes': 24,
        'incometaxes': 26,
        # 'consolidatednetincomeloss': 27, to be computed
        'noncontrollinginterestincome': 29,
        'discontinuedoperations': 34,
        'extraordinaryitems': 35,
        'preferredstockdividendsdeclared': 38,
        'netincomelossattributablecommonshareholders': 39,
        'dilutedsharesoutstanding': 41,
        'commonstockdividendsdeclared': 42,
    },
    'balance_sheet_statement': {
        'cashequivalents': 47,
        'shortterminvestments': 48,
        'tradereceivables': 51,
        'financingreceivables': 52,
        'otherreceivables': 53,
        'inventories': 55,
        'prepaidexpensesothercurrentassets': 56,
        'divestmentassetscurrent': 57,
        'deferredtaxassetscurrent': 58,
        'propertyplantequipment': 62,
        'propertyplantequipmentnet': 64,
        'goodwill': 65,
        'intangibleassetsnet': 66,
        'longterminvestments': 67,
        'divestmentassetsnoncurrent': 68,
        'deferredtaxassetsnoncurrent': 69,
        'receivablesnoncurrent': 70,
        'otherassets': 71,
        'accountspayable': 74,
        'accruedothercurrentliabilities': 77,
        'debtcurrent': 79,
        'taxespayablecurrent': 80,
        'deferredliabilitiescurrent': 81,
        'customeradvancesdepositscurrent': 82,
        'divestmentliabilitiescurrent': 83,
        'dividendspayable': 84,
        'employeerelatedliabilitiescurrent': 85,
        'longtermdebtcapitalleaseobligations': 90,
        'deferredtaxesnoncurrent': 93,
        'taxespayablenoncurrent': 94,
        'divestmentliabilitiesnoncurrent': 95,
        'noncontrollinginterest': 96,
        'pensionotherpostretirementliabilities': 97,
        'employeerelatedliabilitiesnoncurrent': 98,
        'othernoncurrentliabilities': 99,
        'preferredstock': 103,
        'commonstock': 104,
        'additionalpaidincapital': 105,
        'retainedearningsdeficit': 106,
        'treasurystock': 107,
        'accumulatedothercomprehensiveincome': 108,
        'commonstockclassasharesoutstanding': 114,
    },
    'cash_flow_statement': {
        'depreciationdepletionamortization': 120,
        'deferredincometaxestaxcredits': 121,
        'sharebasedcompensation': 122,
        'changedeferredrevenue': 123,
        'gainlossonsale': 124,
        'unrealizedgainloss': 125,
        'pensionotherpostretirementbenefits': 126,
        'changetaxespayable': 128,
        'changeinventories': 129,
        'changetradereceivables': 130,
        'changeaccountspayable': 131,
        'changeotheroperatingassetsliabilitiesnet': 132,
        'restructuringimpairmentchargescashflow': 134,
        'changecustomeradvancesdeposits': 135,
        'provisiondoubtfulaccounts': 136,
        'changefinancingreceivables': 137,
        'changeotherreceivables': 138,
        'extraordinaryitemscashflow': 139,
        'otheroperatingactivities': 140,
        'acquisitionsnet': 144,
        'purchasepropertyplantequipment': 145,
        'salepropertyplantequipment': 146,
        'purchaseinvestments': 148,
        'intangiblesnet': 149,
        'salematurityinvestments': 150,
        'otherinvestingactivities': 151,
        'proceedsincentiveplans': 156,
        'cashdividends': 157,
        'dividendsnoncontrollinginterests': 158,
        'noncontrollinginterestsfinancing': 159,
        'equityissuances': 161,
        'equityrepurchases': 162,
        'longtermdebtrepayments': 164,
        'longtermdebtissuances': 165,
        'shorttermdebtissuancesrepayments': 167,
        'otherfinancingactivities': 168,
        'effectcurrencyexchangerate': 171,
        'cashpaidincometaxes': 173,
        'cashpaidinterest': 174,
    },
}


_DIVISORS = defaultdict(lambda: 1000000.0)
_DIVISORS.update({
    'commonstockdividendsdeclared': 1.0,
    'preferredstockdividendsdeclared': 1.0,
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
    if not col:
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


class Cell:

    COORD_RE = re.compile(r'([A-Z]+)(\d+)')

    def __init__(self, coords, odf_cell):
        self.__coords = coords
        self.__cell = odf_cell

    @property
    def coords(self):
        return self.__coords

    @classmethod
    def tupleFromCoords(cls, coords):
        col, row = cls.COORD_RE.match(coords).groups()
        return col, int(row)

    @property
    def column(self):
        return self.tupleFromCoords(self.coords)[0]

    def setValue(self, value, vtype="string", is_formula=False):
        self.__cell.setAttribute("valuetype", vtype)
        if is_formula:
            assert self.__cell.getAttribute('formula'), f"Cell {self.coords} doesn't have formula."
            self.__cell.setAttribute('value', value)
        else:
            if vtype in ("float", "string"):
                self.__cell.setAttribute("value", value)
            else:
                raise ValueError(f"Value type {vtype} not supported")
            if self.__cell.getAttribute("formula"):
                self.__cell.removeAttribute("formula")
        if self.__cell.firstChild is not None:
            self.__cell.removeChild(self.__cell.firstChild)
        self.__cell.addElement(P(text=str(value)))

    def getValue(self):
        value = self.__cell.getAttribute('value') or 0
        if self.__cell.getAttribute("valuetype") in ("float", "currency"):
            try:
                value = float(value)
            except:
                pass
        return value

    def getFormula(self):
        return self.__cell.getAttribute('formula')

    def evalFormula(self, filled_cells):
        formula = self.getFormula()
        return evaluate(formula, filled_cells)


class Row:

    def __init__(self, idx, odf_row):
        self.__rowindex = idx
        self.__odf_row = odf_row

    @property
    def index(self):
        return self.__rowindex

    @staticmethod
    def _copycell(odfcell, **kwargs):
        kw = {}
        for _, attr in odfcell.attributes.keys():
            attr = attr.replace('-', '')
            kw[attr] = odfcell.getAttribute(attr)
        kw.update(kwargs)
        return TableCell(**kw)

    def getCell(self, col):
        colord = _get_column_ord(col)
        coln = 0
        for cell in self.__odf_row.getElementsByType(TableCell):
            n = int(cell.getAttribute('numbercolumnsrepeated') or 1)
            if coln + n > colord:
                if colord - coln == 1:
                    # we are in the target cell, just remove the repeated attribute
                    # and insert a new cell with repeated decreased by 1
                    cell.removeAttribute('numbercolumnsrepeated')
                    newcell = self._copycell(cell, numbercolumnsrepeated=n-1)
                    self.__odf_row.insertBefore(newcell, cell.nextSibling)
                    return Cell(f'{col}{self.index}', cell)
                # we are in a previous cell, just remove the repeated, and add two new cells,
                # one being the target cell and another with repeated decreased by two
                cell.removeAttribute('numbercolumnsrepeated')
                nextSibling = cell.nextSibling
                newcell = self._copycell(cell)
                self.__odf_row.insertBefore(newcell, nextSibling)
                if n > 2:
                    newnewcell = self._copycell(cell)
                    if n > 3:
                        newnewcell.setAttribute('numbercolumnsrepeated', n - 2)
                    self.__odf_row.insertBefore(newnewcell, nextSibling)
                return Cell(f'{col}{self.index}', newcell)
            if coln + n == colord:
                return Cell(f'{col}{self.index}', cell)
            coln += n
        raise ValueError


class keydefaultdict(defaultdict):
    def __missing__(self, key):
        if self.default_factory is None:
            raise KeyError( key )
        else:
            ret = self[key] = self.default_factory(key)
            return ret


class Sheet:

    def __init__(self, odf_sheet):
        self.__sheet = odf_sheet

    def getCell(self, coord):
        col, row = Cell.tupleFromCoords(coord)
        rowel = self._getRowByIndex(row)
        return rowel.getCell(col)

    def _getRowByIndex(self, idx):
        assert idx > 0, f"Wrong row index: {idx}"
        rown = 0
        for row in self.__sheet.getElementsByType(TableRow):
            n = int(row.getAttribute('numberrowsrepeated') or 1)
            if rown + n > idx:
                if idx - rown == 1:
                    # we are in the target row, just remove the repeated attribute
                    # and insert a new row with repeated decreased by 1
                    row.removeAttribute('numberrowsrepeated')
                    newrow = TableRow(stylename=row.getAttribute('stylename'), numberrowsrepeated=n - 1)
                    self.__sheet.insertBefore(newrow, row.nextSibling)
                    return Row(idx, row)
                # we are in a previous row, just remove the repeated, and add two new rows,
                # one being the target row and another with repeated decreased by two
                row.removeAttribute('numberrowsrepeated')
                nextSibling = row.nextSibling
                newrow = TableRow(stylename=row.getAttribute('stylename'))
                self.__sheet.insertBefore(newrow, nextSibling)
                if n > 2:
                    newnewrow = TableRow(stylename=row.getAttribute('stylename'))
                    if n > 3:
                        newnewrow.setAttribute('numberrowsrepeated', n - 2)
                    self.__sheet.insertBefore(newnewrow, nextSibling)
                return Row(idx, newrow)
            if rown + n == idx:
                return Row(idx, row)
            rown += n
        raise ValueError(f"Error retrieving row {idx}")


class Document:
    def __init__(self, filename):
        self.__filename = filename
        self.__doc = load(filename)

    def getSheet(self, name):
        for sheet in self.__doc.spreadsheet.childNodes:
            if sheet.getAttribute('name') == name:
                return Sheet(sheet)
        raise ValueError(f"No sheet with name {name}.")

    def save(self, filename=None):
        self.__doc.save(filename or self.__filename)


class Process:
    def __init__(self):

        parser = argparse.ArgumentParser()
        parser.add_argument('spreadsheet', help='Update given spreadsheet')
        parser.add_argument('ifile', help='Process given input json (generated by spider)')
        parser.add_argument('column', help='Column where to start to add new data.')

        self.args = parser.parse_args()
        self.__filled_cells = None

    def run(self):
        company, statement, period_type = _FILE_RE.match(self.args.ifile).groups()
        doc = Document(self.args.spreadsheet)
        doc.save(self.args.spreadsheet + '.back')
        sheet = doc.getSheet(company)
        self.__filled_cells = keydefaultdict(lambda key: sheet.getCell(key).getValue())
        translations = _TRANSLATION[statement]
        new_data = json.load(open(self.args.ifile))
        column = self.args.column
        for fundamental in new_data['fundamentals'][::-1]:
            if period_type == 'annual' and not fundamental['annual_period']:
                continue
            if period_type != 'annual' and fundamental['annual_period'] and statement != 'balance_sheet_statement':
                continue

            print(f"End period: {fundamental['end_period']}")
            # update header
            cell = sheet.getCell(f'{column}1')
            if period_type == 'annual':
                cell.setValue(str(fundamental['fiscal_year']))
            elif period_type in ('quarter', 'ttm'):
                quarter = fundamental['fiscal_quarter']
                if quarter == 4:
                    cell.setValue(str(fundamental['fiscal_year']))
                else:
                    cell.setValue(f"TTM {fundamental['fiscal_year']}.{'I'*quarter}")

            # update column
            for tag in fundamental['tags']:
                tag['tag'] = tag['tag'].lower()
                if tag['tag'] in translations:
                    value = tag['value'] / _DIVISORS[tag['tag']]
                    if value:
                        row = translations[tag['tag']]
                        cell = sheet.getCell(f'{column}{row}')
                        cell.setValue(value, 'float')
                        print(f"Updated cell {column}{row} with value {value} ({tag['tag']})")

            # update formulas
            for row in range(1, _MAX_NUM_ROWS + 1):
                coord = f'{column}{row}'
                try:
                    self.__filled_cells[coord] = sheet.getCell(coord).getValue()
                except ValueError:
                    pass
            for row in range(1, _MAX_NUM_ROWS + 1):
                try:
                    cell = sheet.getCell(f'{column}{row}')
                except ValueError:
                    pass
                else:
                    formula = cell.getFormula()
                    if formula is not None:
                        try:
                            value = cell.evalFormula(self.__filled_cells)
                            value = evaluate(formula, self.__filled_cells)
                        except Exception as e:
                            print(f'Error evaluating cell {column}{row}: {formula}')
                            print(f'{e!r}')
                            return
                        print(f'Evaluated cell {column}{row} with result {value} ({formula})')
                        self.__filled_cells[f'{column}{row}'] = value
                        cell.setValue(value, is_formula=True)

            column = _incr_column(column)

        doc.save()
        print("Saved", self.args.spreadsheet)


if __name__ == '__main__':
    process = Process()
    process.run()

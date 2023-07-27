#!/usr/bin/env python3

"""Generates a set of mastermind to solve i.e. an Excel-workbook to print out
with both participants' sheets and the stops.
"""

import argparse
import openpyxl
import openpyxl.styles
import random
from functools import reduce, lru_cache

HEADING_PER_SHEET = "Lagblankett"

INTRO_TEXT_PER_SHEET = """\
Det här är lagblanketten för Radiomastermind på Skogsrå.

Uppgiften är att få tag på den rätta raden, rätt kombination av
färger.  Fyll i den rätta raden längst upp.

Regler: Antalet svarta anger hur många av den rätta radens färger
som återfinns på rätt plats. Antalet svarta anger hur många av rätta
radens färger som återfinns på fel plats.

Ledtrådar sänds via radio till lagmedlemmen vid kontrollen för att få
reda på hur många svarta och vita som gäller för den raden."""

STOP_HEADING = """kontroll"""
BLACK_HEADING = """svarta"""
WHITE_HEADING = """vita"""
CLUE_HEADING = """ledtråd"""

HEADING_PER_STOP = "Radiomastermindkontroll nummer"

INTRO_TEXT_PER_STOP = """Detta är en kontroll för Radiomastermind på Skogsrå.

Radiomastermind är en patrullövning med radiokommunikation som ni kan
få göra som en programaktivitet om ni bokar ett aktivitetspass med
radioscouting eller som ni kan prova en kväll."""

ROWS_PER_SHEET = 51

INTRO_ALIGNMENT = openpyxl.styles.Alignment(vertical="top",
                                            wrap_text=True)
CENTER_ALIGNMENT = openpyxl.styles.Alignment(horizontal="center")
STOP_ALIGNMENT = CENTER_ALIGNMENT
CLUE_ALIGNMENT = CENTER_ALIGNMENT
COLOR_ALIGNMENT = CENTER_ALIGNMENT
HEADING_FONT = openpyxl.styles.Font(size=9, bold=True)
SIDE = openpyxl.styles.Side(border_style="thin",
                            color='FF000000')

HEADING_SIDE = openpyxl.styles.Side(border_style="double",
                                    color='FF000000')
HEADER_BORDER = openpyxl.styles.Border(top=HEADING_SIDE,
                                       bottom=HEADING_SIDE,
                                       left=HEADING_SIDE,
                                       right=HEADING_SIDE)
CELL_BORDER = openpyxl.styles.Border(top=SIDE,
                                     bottom=SIDE,
                                     left=SIDE,
                                     right=SIDE)

parser = argparse.ArgumentParser(description="Generate a set of mastermind games.")
parser.add_argument('--sheets', type=int, default=2,
                    help='Number of sheets')
parser.add_argument('--columns', type=int, default=5,
                    help='Number of clues to guess')
parser.add_argument('--colors', type=int, default=8,
                    help='Number of colors to choose from')
parser.add_argument('--stops', type=int, default=13,
                    help='Number of stops to go to on the course')
parser.add_argument('--filename', type=str,
                    help='The file where the result is stored',
                    default="mastermind.xlsx")
parser.add_argument('--debug', '-d', action='store_true',
                    help='Activate trace outputs',
                    default=False)

args = None

def random_line():
    """Generate a random line."""
    line = []
    for _ in range(args.columns):
        line.append(random.randint(1, args.colors))
    return line

class TooManyClues(Exception):
    pass

class NoCombinationsLeft(Exception):
    pass

class Sheet(object):
    def __init__(self):
        """Creates a sheet."""
        self.correct = random_line()
        self.clue_lines = []
        self.clue_answers = []
        combs = self.combinations(self.clue_lines)
        while combs > 1:
            if len(self.clue_lines) >= args.stops:
                raise TooManyClues()
            new_line = random_line()
            new_clue_lines = self.clue_lines + [new_line]
            new_combs = self.combinations(new_clue_lines)
            if new_combs < combs:
                self.clue_lines = new_clue_lines
                self.clue_answers.append(self.answer(new_line))
                combs = new_combs
        if args.debug:
            print("Verified that the sheet is solvable.", len(self.clue_lines), "lines.")
        while len(self.clue_lines) < args.stops:
            new_line = random_line()
            self.clue_lines.append(new_line)
            self.clue_answers.append(self.answer(new_line))

    def answer(self, clue_line, correct=None):
        """Returns a tuple of counts for black and white."""
        if correct == None:
            correct = self.correct
        count_black = 0
        for i in range(args.columns):
            if correct[i] == clue_line[i]:
                count_black += 1
        count_white = 0
        for i in range(args.columns):
            if correct[i] in (clue_line[0:i] + clue_line[i + 1:]):
                count_white += 1
        return count_black, count_white

    def combinations(self, clue_lines):
        reduced_combinations = [set(range(1, args.colors + 1)) for _ in range(args.columns)]
        for clue_line in clue_lines:
            black, white = self.answer(clue_line)
            if black == 0 and white == 0:
                for i in range(args.columns):
                    for j in range(args.columns):
                        if clue_line[j] in reduced_combinations[i]:
                            reduced_combinations[i].remove(clue_line[j])
            elif black == 0:
                for i in range(args.columns):
                    if clue_line[i] in reduced_combinations[i]:
                        reduced_combinations[i].remove(clue_line[i])
        if args.debug:
            print("Combinations left:",
                  reduce((lambda x, y: x * y),
                         [len(s) for s in reduced_combinations]))
        combinations = []
        for x in reduced_combinations[0]:
            combinations.append([x])
        for c in range(1, args.columns):
            old_combinations = combinations
            combinations = []
            for x in reduced_combinations[c]:
                for comb in old_combinations:
                    combinations.append(comb + [x])
        valid_combinations = []
        for comb in combinations:
            for clue_line in clue_lines:
                if self.answer(clue_line, comb) != self.answer(clue_line):
                    continue
            valid_combinations.append(comb)
        if len(valid_combinations) == 0:
            raise NoCombinationsLeft()
        return len(valid_combinations)

    def output(self, ws, sheet_identity, start_row, replacement):
        """Will fill the worksheet from line start_line with the sheet.

        After a common header with explanation, there board is generated
        with each line generated with the clue
        ----------------------------------------------------------------------
        | stop# | color | color | color | ... |      | empty | empty | clue  |
        ----------------------------------------------------------------------

        The amount of color columns depends on the given columns.
        The empty columns are to enter "black" and "white" responses.
        Headers are created.
        """
        stop_column = 1
        black_column = 1 + args.columns + 2
        white_column = black_column + 1
        clue_column = white_column + 1

        ws.merge_cells(start_row=start_row, end_row=start_row,
                       start_column=1, end_column=9)
        ws.cell(row=start_row, column=1).value = HEADING_PER_SHEET + " " + sheet_identity
        ws.cell(row=start_row, column=1).alignment = INTRO_ALIGNMENT

        row = start_row + 3
        ws.merge_cells(start_row=row, end_row=row + 12,
                       start_column=1, end_column=9)
        ws.cell(row=row, column=1).value = INTRO_TEXT_PER_SHEET
        ws.cell(row=row, column=1).alignment = INTRO_ALIGNMENT

        row += 15
        for column in range(args.columns):
            ws.cell(row=row, column=2 + column).border = HEADER_BORDER

        row += 2
        ws.cell(row=row, column=stop_column).value = STOP_HEADING
        ws.cell(row=row, column=stop_column).border = HEADER_BORDER
        ws.cell(row=row, column=stop_column).alignment = STOP_ALIGNMENT

        ws.cell(row=row, column=black_column).value = BLACK_HEADING
        ws.cell(row=row, column=black_column).border = HEADER_BORDER
        ws.cell(row=row, column=black_column).alignment = CENTER_ALIGNMENT

        ws.cell(row=row, column=white_column).value = WHITE_HEADING
        ws.cell(row=row, column=white_column).border = HEADER_BORDER
        ws.cell(row=row, column=white_column).alignment = CENTER_ALIGNMENT

        ws.cell(row=row, column=clue_column).value = CLUE_HEADING
        ws.cell(row=row, column=clue_column).border = HEADER_BORDER
        ws.cell(row=row, column=clue_column).alignment = CENTER_ALIGNMENT

        row += 1
        for line in range(args.stops):
            ws.cell(row=row + line, column=stop_column).value = line + 1
            ws.cell(row=row + line, column=stop_column).border = CELL_BORDER
            ws.cell(row=row + line, column=stop_column).alignment = STOP_ALIGNMENT
            
            ws.cell(row=row + line, column=black_column).border = CELL_BORDER
            ws.cell(row=row + line, column=white_column).border = CELL_BORDER

            code = replacement.generate_clue(line + 1, self.clue_answers[line])
            ws.cell(row=row + line, column=clue_column).value = code
            ws.cell(row=row + line, column=clue_column).border = CELL_BORDER

            for column in range(args.columns):
                cell = ws.cell(row=row + line, column=2 + column)
                cell.value = self.clue_lines[line][column]
                cell.border = CELL_BORDER
                cell.alignment = COLOR_ALIGNMENT
                cell.fill = openpyxl.styles.PatternFill("solid", 
                                                        fgColor=openpyxl.styles.Color(indexed=7 + self.clue_lines[line][column]))

        row += line

        assert row - start_row < ROWS_PER_SHEET


class Stops(object):
    """Maintains all informations from all stops as gotten from each sheet.

    Generates clues as information is added acting as the replacement
    object when creating sheets."""
    def __init__(self):
        self.stop_infos = dict()
        self.max_clues = args.stops * args.sheets
        self.next_clue = self.generate_clues()

    def generate_clues(self):
        clues = list(range(100, 100 + self.max_clues))
        random.shuffle(clues)
        for c in clues:
            yield c

    def generate_clue(self, stop, tuple):
        if stop not in self.stop_infos:
            self.stop_infos[stop] = dict()
        if tuple not in self.stop_infos[stop]:
            self.stop_infos[stop][tuple] = next(self.next_clue)
        return self.stop_infos[stop][tuple]

    def output(self, ws, start_row):
        print(self.stop_infos)
        for stop_number, _ in enumerate(range(args.stops), start=1):
            ws.merge_cells(start_row=start_row, end_row=start_row,
                           start_column=1, end_column=9)
            ws.cell(row=start_row, column=1).value = HEADING_PER_STOP + " " + str(stop_number)
            ws.cell(row=start_row, column=1).alignment = INTRO_ALIGNMENT
            row = start_row + 3

            clue_column = 1
            black_column = 2
            white_column = 3

            ws.cell(row=row, column=clue_column).value = CLUE_HEADING
            ws.cell(row=row, column=clue_column).border = HEADER_BORDER
            ws.cell(row=row, column=clue_column).alignment = CLUE_ALIGNMENT
            
            ws.cell(row=row, column=black_column).value = BLACK_HEADING
            ws.cell(row=row, column=black_column).border = HEADER_BORDER
            ws.cell(row=row, column=black_column).alignment = CLUE_ALIGNMENT
            
            ws.cell(row=row, column=white_column).value = WHITE_HEADING
            ws.cell(row=row, column=white_column).border = HEADER_BORDER
            ws.cell(row=row, column=white_column).alignment = CLUE_ALIGNMENT

            row += 2
            for clue, tuple in sorted([(v,k,) for k, v in self.stop_infos[stop_number].items()]):
                blacks, whites = tuple
                ws.cell(row=row, column=clue_column).value = str(clue)
                ws.cell(row=row, column=clue_column).border = CELL_BORDER
                ws.cell(row=row, column=clue_column).alignment = CLUE_ALIGNMENT

                ws.cell(row=row, column=black_column).value = str(blacks)
                ws.cell(row=row, column=black_column).border = CELL_BORDER
                ws.cell(row=row, column=black_column).alignment = CLUE_ALIGNMENT

                ws.cell(row=row, column=white_column).value = str(whites)
                ws.cell(row=row, column=white_column).border = CELL_BORDER
                ws.cell(row=row, column=white_column).alignment = CLUE_ALIGNMENT

                row += 1

            assert row - start_row < ROWS_PER_SHEET
            start_row += ROWS_PER_SHEET

if __name__ == "__main__":
    args = parser.parse_args()

    wb = openpyxl.Workbook()
    ws = wb.active

    stops = Stops()

    row = 1
    for index in range(args.sheets):
        sheet_number = 1 + index
        s = Sheet()
        print(s.correct)
        print(s.clue_lines)

        s.output(ws, str(sheet_number), row, stops)
        row += ROWS_PER_SHEET

    stops.output(ws, row)

    wb.save("mm.xlsx")

import math
import pprint
import csv
from openpyxl import load_workbook
from argparse import ArgumentParser

def mean(numbers):
    s = float(sum(numbers))
    print(f"SUM: {s}")
    return s / max(len(numbers), 1)

def trunc_float(f, n):
    return math.floor(f * 10 ** n) / 10 ** n

def calc_handicap(n):
    return math.floor(0.96 * n)

# Takes a sorted list of scores, lowest to highest
# Returns the average of the correct number of low scores
def avg_course_index(lowest_scores):
    if len(lowest_scores) < 7:
        n = 1
    elif len(lowest_scores) < 9:
        n = 2
    elif len(lowest_scores) < 10:
        n = 3
    elif len(lowest_scores) < 12:
        n = 4
    elif len(lowest_scores) < 14:
        n = 5
    elif len(lowest_scores) < 16:
        n = 6
    elif len(lowest_scores) == 17:
        n = 7
    elif len(lowest_scores) == 18:
        n = 8
    elif len(lowest_scores) == 19:
        n = 9
    else:
        n = 10
    return trunc_float(mean(lowest_scores[0:n]), 1)

parser = ArgumentParser()
parser.add_argument("spreadsheet")
args = parser.parse_args()

# Run formulas the first time to get things like diff
wb = load_workbook(filename=args.spreadsheet,
    data_only=True)
players_sheet = wb['Players']

players_data = []
for row in players_sheet.iter_rows():
    players_data.append(row)

player = None
scores = []
data_row = False
header_row = False
score_rows = False
final_players = {}
pp = pprint.PrettyPrinter(indent=4)
for r in players_data:
    # print(f"Row: {[x.value for x in r]}")
    if header_row:
        # print("Header row")
        header_row = False
        score_rows = True
    elif data_row:
        # Pull in player data
        # player_data = [x.value for x in r[4:9]]
        player_data = list(["{}{}".format(x.column, x.row) for x in r[4:9]])
        print(f"Player data: {player_data}")
        data_row = False
        header_row = True
    elif r[4].value == "Handicap":
        # We have a new player
        player = r[1].value
        print(f"New player: {player}")
        # print([x.value for x in r])
        data_row = True
    elif score_rows and r[1].value is None:
        # print(f"Finished scores for {player}")
        # Calculate new values
        sorted_by_score = sorted(scores, key=lambda x: x[4])
        differentials = [x[8] for x in sorted_by_score]
        print(f"Diffs:")
        pp.pprint(differentials)
        ai = avg_course_index(differentials)
        print(f"Avg Course Index: {ai}")
        hc = calc_handicap(ai)
        print(f"Handicap: {hc}")
        # TODO: Save off new values for player
        final_players[player] = {
            player_data[0]: hc,
            player_data[1]: ai
        }
        # Clear data for next player
        player = None
        score_rows = False
        scores = []
    elif score_rows:
        # Scores
        # print(f"Scores: {[x.value for x in r]}")
        scores.append([x.value for x in r])
    else:
        pass
wb.close()
wb = load_workbook(filename=args.spreadsheet)
players_sheet = wb['Players']
for player, data in final_players.items():
    print(player, data)
    for k, v in data.items():
        players_sheet[k] = v
# Save out spreadsheet
wb.save("output.xlsx")   
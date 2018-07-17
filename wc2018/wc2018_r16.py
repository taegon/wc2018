import xlrd
from scipy.stats import rankdata

FINAL_SCORE_FILE = "../input/r16/world_cup_2018r16_final.xlsx"
final_team_list = {}
INPUT_DIR = "../input/r16/"
INPUT_DIR_1ST = "../input/"
PLAYER = {
    "Jennifer": "world_cup_2018r16-JS.xlsx",
    "Lloyd": "world_cup_2018r16LR.xlsx",
    "Mo": "world_cup_2018r16_Mo.xlsx",
    "Nathaniel": "world_cup_2018_Ns.xlsx",
    "Taegon": "world_cup_2018r16_TK.xlsx",
    "Tim": "world_cup_2018r16_Tim.xlsx",
    "Rylie": "world_cup_2018r16_RP.xlsx",
}

PLAYER_1st = {
    "Jennifer": "world_cup_2018_Jen.xlsx",
    "Lloyd": "world_cup_2018LR.xlsx",
    "Mo": "world_cup_2018_Mo.xlsx",
    "Nathaniel": "world_cup_2018_Ns.xlsx",
    "Taegon": "world_cup_2018_Taegon.xlsx",
    "Tim": "world_cup_2018_Tim.xlsx",
    "Rylie": "world_cup_2018_RP.xlsx",
}

PLAYER_RANK = []

ROUND_KEY = (
    "R16", "R8", "R4", "R2", "R1"
)

LINEPLOT = [
    ["Round of 16", 1],
    ["Round of 8", 2],
    ["Semi-final", 3],
    ["Final", 4],
]

player_score = {}

ROUND_16 = (( 9 + 4 * 0, 51), (10 + 4 * 0, 51),
            ( 9 + 4 * 1, 51), (10 + 4 * 1, 51),
            ( 9 + 4 * 2, 51), (10 + 4 * 2, 51),
            ( 9 + 4 * 3, 51), (10 + 4 * 3, 51),
            ( 9 + 4 * 4, 51), (10 + 4 * 4, 51),
            ( 9 + 4 * 5, 51), (10 + 4 * 5, 51),
            ( 9 + 4 * 6, 51), (10 + 4 * 6, 51),
            ( 9 + 4 * 7, 51), (10 + 4 * 7, 51),)

ROUND_8 = (( 11 + 8 * 0, 51 + 6 * 1), (12 + 8 * 0, 51 + 6 * 1),
           ( 11 + 8 * 1, 51 + 6 * 1), (12 + 8 * 1, 51 + 6 * 1),
           ( 11 + 8 * 2, 51 + 6 * 1), (12 + 8 * 2, 51 + 6 * 1),
           ( 11 + 8 * 3, 51 + 6 * 1), (12 + 8 * 3, 51 + 6 * 1),)

ROUND_4 = (( 15, 51 + 6 * 2), (16, 51 + 6 * 2),
           ( 31, 51 + 6 * 2), (32, 51 + 6 * 2),)

# I feel it's duplicate Round of 4
ROUND_2 = (( 22, 51 + 6 * 3), (23, 51 + 6 * 3),)

# Match final score
ROUND_1 = ((40, 51 + 6 * 1 + 9),)


def convert_int(val):
    if val is not None and val != "":
        return int(val)
    return "-"


def read_final_score():
    global final_team_list
    final_team_list = read_round_team(FINAL_SCORE_FILE)


def read_round_team(filename):
    round_team = {}
    team = []
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_name('2018 World Cup')

    team.clear()
    for r, c in ROUND_16:
        team.append((sheet.cell(r, c).value, 5))
    round_team["R16"] = team[:]

    team.clear()
    for r, c in ROUND_8:
        team.append((sheet.cell(r, c).value, 10))
    round_team["R8"] = team[:]

    team.clear()
    for r, c in ROUND_4:
        team.append((sheet.cell(r, c).value, 20))
    round_team["R4"] = team[:]

    team.clear()
    for r, c in ROUND_2:
        team.append((sheet.cell(r, c).value, 40))
    round_team["R2"] = team[:]

    team.clear()
    for r, c in ROUND_1:
        team.append((sheet.cell(r, c).value, 80))
    round_team["R1"] = team[:]

    return round_team


def count_score(player_team_list, start=None, last=None):
    if start is None:
        start = 0
    if last is None:
        last = 5
    player_match = []
    for i in range(start, last):
        sol = final_team_list[ROUND_KEY[i]]
        pred = player_team_list[ROUND_KEY[i]]
        player_match.extend(list(set(sol).intersection(set(pred))))

    player_score = 0
    for team, score in player_match:
        player_score += score

    return player_score


def read_scores():
    for player_name, filename in PLAYER.items():
        score_list = read_round_team(INPUT_DIR + filename)
        player_score[player_name] = dict()
        player_score[player_name]["raw"] = score_list
        player_score[player_name]["score"] = count_score(score_list, 1)

    for player_name, filename in PLAYER_1st.items():
        if player_name not in player_score:
            continue
        score_list = read_round_team(INPUT_DIR_1ST + filename)
        player_score[player_name]["raw_1st"] = score_list
        player_score[player_name]["score_1st"] = count_score(score_list)

    print(player_score)


def wrap_tag(tag, list):
    content = ""
    for title in list:
        if "#" in title:
            v, attr = title.split("#")
            content += "<{} class={}>{}</{}>\n".format(tag, attr, v, tag)
        else:
            content += "<{}>{}</{}>\n".format(tag, title, tag)
    return content


def write_table(f):
    for r_key in ROUND_KEY[1:]:
        final_teams = final_team_list[r_key]
        for i, team_info in enumerate(final_teams):
            f.write("<tr>\n")
            row_content = []
            if i == 0:
                f.write("<td rowspan=\"{}\">{}</td>".format(len(final_teams), r_key))

            row_content.append("{}".format(final_teams[i][0]))
            for p in PLAYER:
                score_attr = "grey"
                if player_score[p]["raw"][r_key][i] in final_teams:
                    score_attr = "correct"
                elif final_teams[i][0].startswith("W") or final_teams[i][0] == "":
                    score_attr = "grey"
                else:
                    score_attr = "wrong"

                row_content.append("{}#{}".format(player_score[p]["raw"][r_key][i][0], score_attr))
            f.write(wrap_tag("td", row_content))
        f.write("</tr>\n")


def write_table_1st(f, end_idx):
    for r_key in ROUND_KEY[:end_idx]:
        final_teams = final_team_list[r_key]
        for i, team_info in enumerate(final_teams):
            f.write("<tr>\n")
            row_content = []
            if i == 0:
                f.write("<td rowspan=\"{}\">{}</td>".format(len(final_teams), r_key))

            row_content.append("{}".format(final_teams[i][0]))
            for p in PLAYER:
                score_attr = "grey"
                if player_score[p]["raw_1st"][r_key][i] in final_teams:
                    score_attr = "correct"
                elif final_teams[i][0].startswith("W"):
                    score_attr = "grey"
                else:
                    score_attr = "wrong"

                row_content.append("{}#{}".format(player_score[p]["raw_1st"][r_key][i][0], score_attr))
            f.write(wrap_tag("td", row_content))
        f.write("</tr>\n")


def round16_winner():
    p_score = []
    for j, p in enumerate(PLAYER):
        p_score.append(count_score(player_score[p]["raw"]))
    max_score = max(p_score)

    winner = []
    for i, p in enumerate(PLAYER):
        if max_score == p_score[i]:
            winner.append(p)

    return "Round of 16's winner: " + "<b>" + ", ".join(winner) + "</b>"


def build_html():
    contents = None
    with open("../input/template_16.html", "r") as f:
        contents = f.readlines()

    with open("../docs/round_16.html", "w", encoding="UTF-8") as f:
        for line in contents:
            if "{{SCORE_TABLE}}" in line:
                f.write("<table class=\"blueTable\">")
                f.write("<thead>")
                header_list = ["Stage#header", "Teams#header"]
                # for p in PLAYER:
                #     header_list.append(p + " (" + str(player_score[p]["score"]) + ")#player")
                #     header_list.append("" + "#score")
                # f.write(wrap_tag("th", header_list))
                header_content = wrap_tag("th", header_list)
                for p in PLAYER:
                    header_content += "<th width=\"150\">{} ({}) <a href=\"{}\" target=\"_blank\">{}</a></th>".format(p, player_score[p]["score"], PLAYER[p], u"\u21E9")
                f.write(header_content)
                f.write("</thead>")
                write_table(f)
                f.write("</table>")
            elif "{{SCORE_TABLE_1ST}}" in line:
                f.write("<table class=\"blueTable\">")
                f.write("<thead>")
                header_list = ["Stage#header", "Teams#header"]
                header_content = wrap_tag("th", header_list)
                for p in PLAYER:
                    header_content += "<th width=\"150\">{} ({}) <a href=\"{}\" target=\"_blank\">{}</a></th>".format(p, player_score[p]["score_1st"], PLAYER_1st[p], u"\u21E9")
                f.write(header_content)
                f.write("</thead>")
                write_table_1st(f, 5)
                f.write("</table>")
            elif "{{ROUND16_WINNER}}" in line:
                lineplot_content = round16_winner()
                line = line.replace("{{ROUND16_WINNER}}", lineplot_content)
                f.write(line)
            else:
                f.write(line)


def main():
    read_final_score()
    read_scores()
    build_html()


if __name__ == "__main__":
    main()
import xlrd
from scipy.stats import rankdata

FINAL_SCORE_FILE = "../input/world_cup_2018_final.xlsx"
final_team_list = []
final_score = []
INPUT_DIR = "../input/"
PLAYER = {
    "Jennifer": "world_cup_2018_Jen.xlsx",
    "Lloyd": "world_cup_2018LR.xlsx",
    "Mo": "world_cup_2018_Mo.xlsx",
    "Nathaniel": "world_cup_2018_Ns.xlsx",
    "Taegon": "world_cup_2018_Taegon.xlsx",
    "Tim": "world_cup_2018_Tim.xlsx",
    "Rylie": "world_cup_2018_RP.xlsx",
}
PLAYER_RANK = []
RANKFLOW = [
    ["June 28 (R3)", 48],
    ["June 27", 44],
    ["June 26", 40],
    # ["June 25", 36],
    ["June 24 (R2)", 32],
    ["June 22", 26],
    ["June 20", 20],
    ["June 19 (R1)", 17],
]
LINEPLOT = [
    ["June 19 (R1)", 17],
    ["June 20", 20],
    ["June 21", 23],
    ["June 22", 26],
    ["June 24 (R2)", 32],
    ["June 25", 36],
    ["June 26", 40],
    ["June 27", 44],
    ["June 28 (R3)", 48],
]

player_score = {}
SKIP_ROW = 14


def convert_int(val):
    if val is not None and val != "":
        return int(val)
    return "-"


def read_final_score():
    global final_score
    final_team_list.clear()
    workbook = xlrd.open_workbook(FINAL_SCORE_FILE)
    sheet = workbook.sheet_by_name('2018 World Cup')
    for r in range(6, 54):
        final_team_list.append([sheet.cell(r, 4).value, sheet.cell(r, 7).value])
    final_score = read_score_list(FINAL_SCORE_FILE)


def read_score_list(filename):
    score_list = []
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_name('2018 World Cup')
    for r in range(6, 54):
        sc_left = convert_int(sheet.cell(r, 5).value)
        sc_right = convert_int(sheet.cell(r, 6).value)
        if sc_left == "-":
            wdl = "-"
        else:
            wdl = "W" if sc_left > sc_right else "L" if sc_left < sc_right else "D"

        score_list.append([sc_left, sc_right, wdl])
    return score_list


def count_score(player_list, last=None):
    if last is None:
        last = len(final_score)
    score = 0
    for i in range(SKIP_ROW, last):
        sol = final_score[i]
        pred = player_list[i]
        if sol[0] == pred[0] and sol[1] == pred[1]:
            score += 2
        if sol[2] == pred[2]:
            score += 1
    return score


def read_scores():
    for player_name, filename in PLAYER.items():
        score_list = read_score_list(INPUT_DIR + filename)
        player_score[player_name] = dict()
        player_score[player_name]["raw"] = score_list
        player_score[player_name]["score"] = count_score(score_list)

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
    for i in range(len(final_team_list)):
        f.write("<tr>\n")
        row_content = []
        row_content.append("{} - {}".format(final_team_list[i][0], final_team_list[i][1]))
        row_content.append("{}:{} ({})".format(final_score[i][0], final_score[i][1], final_score[i][2]))
        header_list = ["Teams#header", "Score#header"]
        for p in PLAYER:
            score_attr = "grey"
            wdl_attr = "grey"
            if i >= SKIP_ROW:
                if player_score[p]["raw"][i][0] == final_score[i][0] and player_score[p]["raw"][i][1] == final_score[i][1]:
                    score_attr = "correct"
                elif final_score[i][2] != "-":
                    score_attr = "wrong"
                if player_score[p]["raw"][i][2] == final_score[i][2]:
                    wdl_attr = "correct"
                elif final_score[i][2] != "-":
                    wdl_attr = "wrong"

            row_content.append("{}:{}#{}".format(player_score[p]["raw"][i][0], player_score[p]["raw"][i][1], score_attr))
            row_content.append("{}#{}".format(player_score[p]["raw"][i][2], wdl_attr))
        f.write(wrap_tag("td", row_content))
        f.write("</tr>\n")


def make_rank_content():
    rank_size = len(RANKFLOW)
    PLAYER_RANK.clear()

    player_ranks = []
    player_scores = []
    for i in range(rank_size):
        p_score = []
        for j, p in enumerate(PLAYER):
            p_score.append(count_score(player_score[p]["raw"], RANKFLOW[i][1]))
        p_rank = rankdata(p_score, method='ordinal')
        p_rank = len(p_rank) - p_rank + 1
        player_scores.append(p_score)
        player_ranks.append(p_rank)
    print(player_scores)
    print(player_ranks)

    p_score = []
    for j, p in enumerate(PLAYER):
        p_score.append(count_score(player_score[p]["raw"]))
    final_rank = rankdata(p_score, method='ordinal')
    final_rank = len(final_rank) - final_rank + 1

    rank_series_text = ""

    for i, p in enumerate(PLAYER):
        template_str = """
        {{
          "text":"{}",
          "ranks":[{}],
          "rank":{}
        }},
        """

        rank_series_text += template_str.format(p, ",".join([str(x[i]) for x in player_ranks]), final_rank[i])
        PLAYER_RANK.append([p, final_rank[i], i])

    return rank_series_text


def make_lineplot_content():
    lineplot_size = len(LINEPLOT)

    player_scores = []
    for i in range(lineplot_size):
        p_score = []
        for j, p in enumerate(PLAYER):
            p_score.append(count_score(player_score[p]["raw"], LINEPLOT[i][1]))
        player_scores.append(p_score)
    print(player_scores)


    lineplot_series_text = ""
    # {"values": [20, 40, 25, 50, 15, 45, 33, 34]},
    # {"values": [5, 30, 21, 18, 59, 50, 28, 33]}
    PLAYER_RANK.sort(key=lambda x: x[1])
    for player in PLAYER_RANK:
        p = player[0]
        i = player[2]
        template_str = """
        {{
          "values":[{}],
          "text": '{}',
        }},
        """

        lineplot_series_text += template_str.format(",".join([str(x[i]) for x in player_scores]), p)

    return lineplot_series_text


def group_stage_winner():
    p_score = []
    for j, p in enumerate(PLAYER):
        p_score.append(count_score(player_score[p]["raw"]))
    max_score = max(p_score)

    winner = []
    for i, p in enumerate(PLAYER):
        if max_score == p_score[i]:
            winner.append(p)

    return "Group Stage winner: " + "<b>" + ", ".join(winner) + "</b>"


def build_html():
    contents = None
    with open("../input/template.html", "r") as f:
        contents = f.readlines()

    with open("../docs/index.html", "w", encoding="UTF-8") as f:
        for line in contents:
            if line.strip() == "{{SCORE_TABLE}}":
                f.write("<table class=\"blueTable\">")
                f.write("<thead>")
                header_list = ["Teams#header", "Score#header"]
                # for p in PLAYER:
                #     header_list.append(p + " (" + str(player_score[p]["score"]) + ")#player")
                #     header_list.append("" + "#score")
                # f.write(wrap_tag("th", header_list))
                header_content = wrap_tag("th", header_list)
                for p in PLAYER:
                    header_content += "<th width=\"70\" colspan=\"2\">{} ({}) <a href=\"{}\" target=\"_blank\">{}</a></th>".format(p, player_score[p]["score"], PLAYER[p], u"\u21E9")
                f.write(header_content)
                f.write("</thead>")
                write_table(f)
                f.write("</table>")
            elif "{{RANKFLOW_LABELS}}" in line:
                rank_label_text = ",".join(["\"{}\"".format(x[0]) for x in RANKFLOW])
                line = line.replace("{{RANKFLOW_LABELS}}", rank_label_text)
                f.write(line)
            elif "{{RANKFLOW_VALUES}}" in line:
                rank_value_text = ",".join(["\"{}\"".format(x[0]) for x in RANKFLOW])
                line = line.replace("{{RANKFLOW_VALUES}}", rank_value_text)
                f.write(line)
            elif "{{RANKFLOW_SERIES}}" in line:
                # {
                #     "text": "Air West",
                #     "ranks": [3, 4, 1],
                #     "rank": 3
                # },
                rank_content = """
                    {
      "text":"Air West",
      "ranks":[3,4,1],
      "rank":3
    },
    {
      "text":"Braniff",
      "ranks":[1,1,5],
      "rank":1
    },
    {
      "text":"Capital",
      "ranks":[5,2,4],
      "rank":4
    },
    {
      "text":"Eastern",
      "ranks":[2,3,2],
      "rank":2
    },
    {
      "text":"Galaxy",
      "ranks":[4,5,3],
      "rank":5
    }
                """
                rank_content = make_rank_content()
                line = line.replace("{{RANKFLOW_SERIES}}", rank_content)
                f.write(line)
            elif "{{LINEPLOT_LABEL}}" in line:
                lineplot_label_text = ",".join(["\"{}\"".format(x[0]) for x in LINEPLOT])
                line = line.replace("{{LINEPLOT_LABEL}}", lineplot_label_text)
                f.write(line)
            elif "{{LINEPLOT_VALUE}}" in line:
                lineplot_content = make_lineplot_content()
                line = line.replace("{{LINEPLOT_VALUE}}", lineplot_content)
                f.write(line)
            elif "{{GROUP_STAGE_WINNER}}" in line:
                lineplot_content = group_stage_winner()
                line = line.replace("{{GROUP_STAGE_WINNER}}", lineplot_content)
                f.write(line)
            else:
                f.write(line)


def main():
    read_final_score()
    read_scores()
    build_html()


if __name__ == "__main__":
    main()
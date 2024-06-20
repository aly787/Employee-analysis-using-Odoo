from openpyxl import load_workbook

wb = load_workbook("Task (project.task) (11).xlsx")
sheet = wb["Sheet1"]
current_company = "Bob"
team = {
    "fig": [],
    "aza": [],
    "yaat": [],
    "maha": [],
    "moah": [],
    "mous": [],
    "saal": [],
    "abr": [],
    "amu": [],
    "maka": [],
    "keh": []
}

for row in sheet.rows:
    if row[0].value is not None:
        rating = row[6].value
        assignee = row[3].value
        if (assignee is None):
            assignee = "None"
        if "fig" in assignee:
            team["fig"].append(rating)
        if "aza" in assignee:
            team["aza"].append(rating)
        if "yaat" in assignee:
            team["yaat"].append(rating)
        if "maha" in assignee:
            team["maha"].append(rating)
        if "moah" in assignee:
            team["moah"].append(rating)
        if "mous" in assignee:
            team["mous"].append(rating)
        if "saal" in assignee:
            team["saal"].append(rating)
        if "abr" in assignee:
            team["abr"].append(rating)
        if "amu" in assignee:
            team["amu"].append(rating)
        if "maka" in assignee:
            team["maka"].append(rating)
        if "keh" in assignee:
            team["keh"].append(rating)
print(team)
#Other Tests:
    # original = row[6].value
    # original1 = row[1].value
    # original2 = row[2].value
    # original3 = row[3].value

    # if row[0].value is not None:
    #     current_owner = row[1].value

    #     if original1 is None:
    #         current_owner = "none"
    #     if original3 is None:
    #         current_subtasks = 0
    #     if original2 is None:
    #         current_assignee = "none"
    #     current_assignee = row[2].value
    #     current_company = row[0].value
    #     current_subtasks = row[3].value

    # if original is None:
    #     continue
    
    # if ("industry" in row[6].value):
    #     counter += 1
    

    # #String here is the comparison value to check what is in the row at column N
    #     print(current_company + "->" + current_owner + "->" + current_assignee + "->    " + str(current_subtasks))

    
from openpyxl import load_workbook

wb = load_workbook("Task (project.task) (8).xlsx")
sheet = wb["Sheet1"]
current_company = "Bob"
fig = 0
aza = 0
yaat = 0
maha  = 0
moah  = 0
mous  = 0
saal = 0
abr = 0
amu = 0
maka = 0
keh = 0
subtasks = 0
assignee = "Bob marley"
company = []

for row in sheet.rows:
    if row[0].value is not None:
        current_company = row[0].value
        subtasks = row[3].value
        assignee = row[2].value
        if (assignee is None):
            assignee = "None"
        if "fig" in assignee:
            fig += 1
        if "aza" in assignee:
            aza += 1
        if "yaat" in assignee:
            yaat += 1
        if "maha" in assignee:
            maha += 1
        if "moah" in assignee:
            moah += 1
        if "mous" in assignee:
            mous += 1
        if "saal" in assignee:
            saal += 1
        if "abr" in assignee:
            abr += 1
        if "amu" in assignee:
            amu += 1
        if "maka" in assignee:
            maka += 1
        if "keh" in assignee:
            keh += 1

    continue
print("fig:" + str(fig) + "--" + "aza:" + str(aza) + "--" + "yaat:" + str(yaat) + "--" + "maha:" + str(maha) + "--" + "moah:" + str(moah) + "--" + "mous:" + str(mous) + "--" + "saal:" + str(saal) + "--" + "abr:" + str(abr) + "--" + "amu:" + str(amu) + "--" + "maka:" + str(maka) + "--" + "keh:" + str(keh))
    #The list will have rows of the project name and the percentage of the purchased hours
    #to be able to 
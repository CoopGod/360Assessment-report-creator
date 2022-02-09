import csv

# parse through all relations and add their stats
def parse(scheduleArg, relationRows, headerRows, relationCount):
    with open(scheduleArg, "r", encoding="utf8") as schedule:
        reader = csv.reader(schedule)
        rowCount = 0
        skills = {}
        importance = {}
        for row in reader:
            if rowCount in relationRows:
                # change relation name to described documentation
                relation = "Employee"
                # if column is skill (based on number beginning the question) assign it to its list. If question starts
                # with opera... then assign it to importance
                itemCount = 0
                previousHeader = None
                for item in row:
                    currentHeader = headerRows[itemCount]
                    if currentHeader[:1].isdigit():
                        try:
                            currentTotal = int(skills[currentHeader[:2]])
                            currentTotal += int(item)
                            skills[currentHeader[:2]] = str(currentTotal)
                        except:
                            skills[currentHeader[:2]] = item
                        previousHeader = currentHeader
                    elif currentHeader[:2] == "Op":
                        try:
                            currentTotal = int(importance[previousHeader[:2]])
                            currentTotal += int(item)
                            importance[previousHeader[:2]] = str(currentTotal)
                        except:
                            importance[previousHeader[:2]] = item
                        previousHeader = currentHeader
                    itemCount += 1
            rowCount += 1
        # use number of employees to create averages for each question
        for key, value in skills.items():
            currentTotal = int(value) / relationCount
            skills[key] = currentTotal
        for key, value in importance.items():
            currentTotal = int(value) / relationCount
            importance[key] = currentTotal
        schedule.close()
        return [relation, skills, importance]
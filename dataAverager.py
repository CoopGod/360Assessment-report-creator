import csv
import sys
import pandas as pd
import random
from io import BytesIO
import matplotlib.pyplot as plt
from PIL import Image
import numpy as np
import docx
from docx.shared import Inches


class Review:
    def __init__(self, focusRelation, skillList, importanceList):
        self.focusRelation = focusRelation
        self.skillList = skillList,
        self.importanceList = importanceList    


def main():
    # Ensure proper CMD Line Arg
    if len(sys.argv) > 2:
        print("Error!")
        return 1
    assessments = dataParse()
    df = createDataFrame(assessments)
    plotData(df)


# function to create Review object to then be input into dataframe
def dataParse():
    # Upload CSV files
    scheduleArg = sys.argv[1]
    # open CSV file and run through data. adding to nested list
    assessments = []
    with open(scheduleArg, "r", encoding="utf8") as schedule:
        reader = csv.reader(schedule)
        # Move through rows including header
        rowCount = 0
        employeeRows = []
        employeeCount = 0
        selfCount = 0
        selfRows = []
        headerList = []
        supervisorCount = 0
        supervisorList = []
        for row in reader:
                # if not header move through and assign based on documentation
                if rowCount != 0:
                    # priliminary setup
                    name = row[1]
                    focus = row[3]
                    # If employee, bring to different loop and average those scores
                    if row[4] == "Peer" or row[4] == "Employee":
                        employeeCount += 1
                        relation = "Employee"
                        employeeRows.append(rowCount)
                    # else add the rest to their own columns
                    elif row[4] == "Self":
                        selfCount += 1
                        selfRows.append(rowCount)
                    else:
                        supervisorCount += 1
                        supervisorList.append(rowCount)
                    rowCount += 1
                else:
                    headerList = row
                    rowCount += 1
    schedule.close()

    # parse through all employees and add their stats
    with open(scheduleArg, "r", encoding="utf8") as schedule:
        reader = csv.reader(schedule)
        rowCount = 0
        skills = {}
        importance = {}
        for row in reader:
            if rowCount in employeeRows:
                # change relation name to described documentation
                relation = "Employee"
                # if column is skill (based on number beginning the question) assign it to its list. If question starts
                # with opera... then assign it to importance
                itemCount = 0
                previousHeader = None
                for item in row:
                    currentHeader = headerList[itemCount]
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
            currentTotal = int(value) / employeeCount
            skills[key] = currentTotal
        for key, value in importance.items():
            currentTotal = int(value) / employeeCount
            importance[key] = currentTotal
        # create Review object for later use 
        assessments.append(Review(relation, skills, importance))
    schedule.close()

    # parse through all Self reviews and add their stats
    with open(scheduleArg, "r", encoding="utf8") as schedule:
        reader = csv.reader(schedule)
        rowCount = 0
        skills = {}
        importance = {}
        for row in reader:
            if rowCount in selfRows:
                # change relation name to described documentation
                relation = "Self"
                # if column is skill (based on number beginning the question) assign it to its list. If question starts
                # with opera... then assign it to importance
                itemCount = 0
                previousHeader = None
                for item in row:
                    currentHeader = headerList[itemCount]
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
            currentTotal = int(value) / selfCount
            skills[key] = currentTotal
        for key, value in importance.items():
            currentTotal = int(value) / selfCount
            importance[key] = currentTotal
        # create Review object for later use 
        assessments.append(Review(relation, skills, importance))
    schedule.close()
            
    # parse through all supervisors and add their stats
    with open(scheduleArg, "r", encoding="utf8") as schedule:
        reader = csv.reader(schedule)
        rowCount = 0
        skills = {}
        importance = {}
        for row in reader:
            if rowCount in supervisorList:
                # change relation name to described documentation
                relation = "Supervisor"
                # if column is skill (based on number beginning the question) assign it to its list. If question starts
                # with opera... then assign it to importance
                itemCount = 0
                previousHeader = None
                for item in row:
                    currentHeader = headerList[itemCount]
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
            currentTotal = int(value) / supervisorCount
            skills[key] = currentTotal
        for key, value in importance.items():
            currentTotal = int(value) / supervisorCount
            importance[key] = currentTotal
        # create Review object for later use 
        assessments.append(Review(relation, skills, importance))
    schedule.close()

    return assessments


# function to create data frame regarding person of focus
def createDataFrame(assessments):
    # init all columns in data frame
    columnList = ['1.','2.','3.','4.','5.','6.','7.','8.','9.','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','35','36','37','38']
    df = pd.DataFrame(columns=columnList)
    # Log all Skill Ratings!
    skillReviewNum = 0
    axisRelations = {}
    for review in assessments:
        skillValues = review.skillList
        # add all review data to df
        temp = pd.DataFrame(skillValues, columns=columnList)
        df = df.append(temp, ignore_index=True)
        # change axis names to reviewers relationship and indicate which they are rating (skill vs importance)
        axisRelations[skillReviewNum] = review.focusRelation + "-SKL"
        skillReviewNum += 1

    # log all Importance Ratings!
    impReviewNum = 0
    for review in assessments:
        impValues = review.importanceList
        # add all review data to df
        temp = pd.DataFrame([impValues], columns=columnList)
        df = df.append(temp, ignore_index=True)
        # change axis names to reviewers relationship and indicate which they are rating (skill vs importance)
        axisRelations[impReviewNum + skillReviewNum] = review.focusRelation + "-IMP"
        impReviewNum += 1
    # rename and sort alphebetically
    df = df.rename(index=axisRelations)
    df = df.sort_index(ascending=False)
    df.plot.bar(y="1.")
    print(df)
    return df

# function to plot dataframe data to report
def plotData(df):
    # get different colors
    colorDictionary = {0: '#fc0403',2: '#008000',4: '#0000FF',6: '#FF7F50',8: '#FFFACD',10: '#FF69B4', 12: '#800080'}
    cmap = []
    bars = len(df.index)
    for num in range(bars):
        if num % 2 != 0:
            cmap.append(cmap[num - 1])
        else:
            cmap.append(colorDictionary[num])
    # save file in report
    # document setup
    document = docx.Document()
    document.add_heading(f'360 Assessment - Averages',0)
    # find complete questions in header of original CSV
    schedule = sys.argv[1]
    with open(schedule, "r", encoding="utf8") as schedule:
        reader = csv.DictReader(schedule)
        fieldnames = reader.fieldnames
        headers = []
        for field in fieldnames:
            field = field.replace("Skilled:", "")
            headers.append(field)
    # create plot and add image of it to docx
    for i in range(1, len(df.columns) + 1):
        iteration = str(i) + "." # to skirt the "." issue with 1-9
        ax = df.plot.bar(y=iteration[:2], color=cmap) # fuck ya
        # pretty graphs <3
        ax.set_ylim(ymax=5.25)
        ax.get_legend().remove()
        plt.yticks(rotation= 90)
        # add image
        plt.savefig('temp.jpg', bbox_inches='tight')
        imageFile = Image.open('temp.jpg')
        # add whitespace to image so rotating isn't wonky
        background = Image.new("RGB", (550, 550), (255,255,255))
        background.paste(imageFile)
        background = background.rotate(270)
        background.save('temp.jpg')
        document.add_picture('temp.jpg', width=Inches(3))
        # add question info/text to document
        paragraph = document.add_paragraph(headers[3+(i*2)])
        # save that sweet, sweet memory
        imageFile.close()
        plt.close()
    # save file under person of focus' name
    document.save('averages-report.docx')


if __name__ == "__main__":
    main()
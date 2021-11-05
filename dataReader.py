import csv
import sys
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
from PIL import Image
import numpy as np
import docx
from docx.shared import Inches


class Review:
    def __init__(self, focusName, focusRelation, reviewName, skillList, importanceList):
        self.focusName = focusName
        self.focusRelation = focusRelation
        self.reviewName = reviewName
        self.skillList = skillList,
        self.importanceList = importanceList   


def main():
    # Ensure proper CMD Line Arg
    if len(sys.argv) > 3:
        print("Error!")
        return 1
    revieweesList = reviewees()
    for review in revieweesList:
        assessments = dataParse(review)
        df = createDataFrame(assessments)
        plotData(df, review)


# function to find all people of focus
def reviewees():
    schedule = sys.argv[1]
    with open(schedule, "r", encoding="utf8") as schedule:
        reader = csv.reader(schedule)
        reviewees = []
        for row in reader:
            if row[3] not in reviewees:
                reviewees.append(row[3])
    # remove header
    reviewees.pop(0)
    return reviewees


# function to create Review object regarding person of focus
def dataParse(reviewee):
    # Upload CSV files
    schedule = sys.argv[1]
    # open CSV file and run through data. adding to nested list
    assessments = []
    with open(schedule, "r", encoding="utf8") as schedule:
        reader = csv.reader(schedule)
        # Move through rows including header
        rowCount = 0
        employeeCount = 0
        for row in reader:
            # if person of focus
            if row[3] == reviewee or rowCount == 0:
                # if not header move through and assign based on documentation
                if rowCount != 0:
                    # priliminary setup (should be changed later for sake of variability)
                    name = row[1]
                    focus = row[3]
                    # change relation name to described documentation (input by customer)
                    employeeCount += 1
                    if row[4] == "Peer" or row[4] == "Employee":
                        relation = f"Employee {employeeCount}"
                    else:
                        relation = row[4]
                    skills = []
                    importance = []
                    # if column is skill (based on number beginning the question) assign it to its list. If question starts
                    # with opera... then assign it to importance
                    itemCount = 0
                    previousHeader = None
                    for item in row:
                        currentHeader = headerList[itemCount]
                        if currentHeader[:1].isdigit():
                            skills.append(f"{currentHeader[:2]}:{item}") # 1-9 remove . later
                            previousHeader = currentHeader
                        elif currentHeader[:2] == "Op":
                            importance.append(f"{previousHeader[:2]}:{item}") # 1-9 remove . later
                        itemCount += 1
                    # create Review object for later use
                    assessments.append(Review(focus, relation, name, skills, importance))
                    rowCount += 1
                else:
                    headerList = row
                    rowCount += 1
    return assessments


# function to create data frame regarding person of focus
def createDataFrame(assessments):
    # init all columns in data frame
    df = pd.DataFrame(columns=['1.','2.','3.','4.','5.','6.','7.','8.','9.','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','35','36','37','38'])
    # Log all Skill Ratings!
    skillReviewNum = 0
    axisRelations = {}
    for review in assessments:
        skillValues = {}
        for skills in review.skillList:
            for skill in skills:
                # move through each skill and append it to the corresponding data frame column
                data = int(skill[3:])
                header = str(skill[:2])
                skillValues[header] = data
        # add all review data to df
        df = df.append(skillValues, ignore_index=True)
        # change axis names to reviewers relationship and indicate which they are rating (skill vs importance)
        axisRelations[skillReviewNum] = review.focusRelation + "-SKL"
        skillReviewNum += 1

    # log all Importance Ratings!
    impReviewNum = 0
    for review in assessments:
        impValues = {}
        for imps in review.importanceList:
            # move through each importance rating and append it to the corresponding data frame column
            data = int(imps[3:])
            header = str(imps[:2])
            impValues[header] = data
        # add all review data to df
        df = df.append(impValues, ignore_index=True)
        # change axis names to reviewers relationship and indicate which they are rating (skill vs importance)
        axisRelations[impReviewNum + skillReviewNum] = review.focusRelation + "-IMP"
        impReviewNum += 1
    df = df.rename(index=axisRelations) 
    return df
        

# function to upload/display graphs using data frame
def plotData(df, reviewee):
    # get different colors 
    cmap = []
    barsHalf = len(df.index) / 2
    bars = len(df.index)
    for num in range(bars):
        if num < barsHalf:
            cmap.append('#b8eff6')
        else:
            cmap.append('#e81a00')
    # save file in report
    # document setup
    document = docx.Document()
    document.add_heading(f'360 Assessment - {reviewee}',0)
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
    document.save(f'{reviewee}-report.docx')


if __name__ == "__main__":
    main()
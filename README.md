# 360-Assessment Report Creator
### Purpose:
Made for [McNeillHR](https://www.legacyhr.ca/), the goal of this **python** application is to quickly move through **data** collected via a specific **Google Form**(please note that this google form is referenced many times but is not available to the public.) and turn all the data into easy to view, visually pleasing **reports** for each person of focus. These reports can then be used by HR to analyze trends and coach the persons of focus in areas where more attention is needed.
### Understanding:

This application takes in a CSV file from the Assessment's Google Form, then finds all the people of focus, or in other words, people who are receiving a report, via the *reviewees* function.
Then the application parses through the raw CSV data and turning it into a more useful form via the *dataParse* function (which also creates a legend for HR to view). 
This refined information is then simply put into a data frame in the *createDataFrame* function. 
The data frame is then utilized by the *plotData* function, to both plot the data on a graph using **Pandas** and **MatPlotLib**. Then it is added, along with relating text, to a report (.docx) to be shipped off to its recipient by HR.

The **dataAverager.py** has a very similar structure, however it combines all relations and all people of focus into one graph. Typically for HR to look at trends and significant outliers. It also comes with a CSV file for the clarification of specific numbers.
### Example Graph
![temp](https://user-images.githubusercontent.com/57197353/141396685-33279d9a-a76d-4720-9b47-c6892ee5219b.jpg)

for full reports, please check out the 'example reports' folder! It contains reports for ~~Individuals and a coresponding legend, as well as~~ an overall graph that combines all answers based on the interviewee's role!

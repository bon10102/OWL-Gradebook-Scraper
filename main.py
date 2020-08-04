from selenium import webdriver
from bs4 import BeautifulSoup
from styleframe import StyleFrame, Styler, utils
import re
import pandas as pd
from os import path
import sys

#sums the contents of an array of numbers
#@param arr the array to be summed
#@return sum the sum of the array
def arrSum(arr):
    sum = 0
    for i in range(0, len(arr)):
        if isinstance(arr[i], (float, int)) == True:
            sum += arr[i]
    return sum

#rounds the content of an array of floats and convert contents to percent
#@param arr the array to be rounded and converted to percent
#@param places the number of places to be rounded
#@return roundedArr the rounded array
def roundArr(arr, places):
    roundedArr = []
    for i in range(0, len(arr)):
        if isinstance(arr[i], float) == True:
            roundedArr.append(round(arr[i], places)*100)
        else:
            roundedArr.append(arr[i])
    return roundedArr

#add path to executable runtime environment for chromedriver
bundle_dir = getattr(sys, "_MEIPASS", path.abspath(path.dirname(__file__)))
path_to_dat = path.join(bundle_dir, "chromedriver.exe")

#choose browser and navigate to webpage
driver = webdriver.Chrome(r"assets\chromedriver.exe")
driver.maximize_window()
driver.get("https://owl.uwo.ca/portal")

#find log-in elements
userId = driver.find_element_by_id("eid")
userPass = driver.find_element_by_id("pw")
submitButton = driver.find_element_by_id("submit")


#log in to OWL
#prompts user to re-enter login information if incorrect
while True:
    userName = input("Please enter your OWL Username: ")
    password = input("Please enter your password: ")
    userId.send_keys(userName)
    userPass.send_keys(password)
    submitButton.click()
    content = driver.page_source
    soup = BeautifulSoup(content, features = "html.parser")
    if soup.find(id = "eid") != None:
        print("Invalid login, please try again")
        userId = driver.find_element_by_id("eid")
        userPass = driver.find_element_by_id("pw")
        submitButton = driver.find_element_by_id("submit")
        userId.clear()
        userPass.clear()
    else:
        break
        
#determine how many courses are in favorites bar
courseButtons = driver.find_elements_by_class_name("link-container")

#extract names of all courses in favorites
courseNames = []
for a in soup.findAll("a", class_ = "link-container"):
    name = re.sub(r"\n", "", a.get_text())
    courseNames.append(name)
courseNames.pop(0)

#array with StyleFrame objects, stores an excel sheet for each course
sf = [None] * len(courseNames)
finalGrades = []

#navigate to the gradebook in each course
for i in range(len(courseNames)):
    courseButtons = driver.find_elements_by_class_name("link-container")
    courseButtons.pop(0)
    courseButtons[i].click()
    gradebook = driver.find_elements_by_link_text("Gradebook")

    #navigate to the gradebook page if it exits
    if len(gradebook) > 0:
        #scroll to gradebook so it is visible in the viewport to avoid clicking on wrong element
        driver.execute_script("arguments[0].scrollIntoView();", gradebook[0])
        gradebook[0].click()
        content = driver.page_source
        soup = BeautifulSoup(content, features = "html.parser")

        #arrays to store category data
        categoryNames = []
        categoryMarks = []
        categoryWeight = []
        categoryWeightGrade = []
        #arrays to store data in the excel sheet output
        names = []
        grades = []
        marks = []
        maxMark = []
        weight = []
        weightedGradeCalc = []
        #stores number of assignments with no grade data and with the "not included in weight calculation" flag
        #used in assignment weight calculation to re-distribute weight to assignments with grades
        noGradeData = 0
        #stores the number of categories in each gradebook.
        #used to order elements correctly in excel sheet output arrays, as well as assignment weight calculation
        numberOfCats = 0

        #find the gradebook table class
        for gbtable in soup.findAll("div", class_ = "gb-summary-grade-panel"):

            #find all the assessment categories
            for allCategories in gbtable.findAll("tbody", class_ = "gb-summary-category-tbody"):
                #search for desired elements
                category = allCategories.find("span", class_ = "gb-summary-category-name").get_text()
                catGrade = allCategories.find("td", class_ = "gb-summary-category-grade").get_text()
                catWeight =  allCategories.find("td", class_ = "gb-summary-category-weight weight-col").get_text()
                categoryNames.append(category)
                categoryMarks.append(catGrade)

                #check if there is a grade and weight under the category to avoid cast errors
                if catWeight != "":
                    #remove '%' character after grade weight and convert to fraction 
                    catWeight = float(re.sub(r"%", "", catWeight)) * 0.01 
                elif catGrade == "-":
                    catWeight = 0
                #if no category weight is given, ask user to manually input category weight
                else:
                    while True:
                        assnWeight = input("Category weight for " + category + " in " + courseNames[i] + " was not provided. Please enter it now (in percent)\n")
                        if catWeight.isdigit() == False:
                            print("Invalid input, please try again")
                            catWeight = input()
                        elif float(catWeight) < 0 or float(catWeight) > 100:
                            print("Inputs must be between 0% - 100%, please try again")
                            catWeight = input()
                        else:
                            catWeight = float(catWeight) * 0.01
                            break
                categoryWeight.append(catWeight)

            #find all assessments in the different categories
            #some gradebooks have no assignment categories. Check if this is the case
            if len(categoryNames) >= 1:

                #search for desired elements
                for allAssessments in gbtable.findAll("tbody", class_ = "gb-summary-assignments-tbody"):
                    names.append(categoryNames[numberOfCats])
                    marks.append(categoryMarks[numberOfCats])
                    weight.append(categoryWeight[numberOfCats])
                    weightedGradeCalc.append("")
                    maxMark.append("100%")
                    numberOfCats += 1
                    noGradeData = 0

                    #find assessments in each category with no grade data
                    for noGradeAssessments in allAssessments.findAll("tr", class_ = re.compile(r"^\bgb-summary-grade-row (odd|even)\b$")):
                        noWeightMarkers = noGradeAssessments.find("span", class_ = "gb-flag-not-counted")
                        assnGrade = noGradeAssessments.find("span", class_ = "gb-summary-grade-score-raw").get_text()
                        if assnGrade == "" or noWeightMarkers != None:
                            noGradeData = noGradeData + 1
                
                    #find individual assessments in each category
                    for assessments in allAssessments.findAll("tr", class_ = re.compile(r"^\bgb-summary-grade-row (odd|even)\b$")):
                        #search for desired elements
                        assessment = assessments.find("span", class_ = "gb-summary-grade-title").get_text()
                        assnGrade = assessments.find("span", class_ = "gb-summary-grade-score-raw").get_text()
                        assnGradeAll = assessments.findAll("span", class_ = "gb-summary-grade-score-raw")
                        assnMaxGrade = assessments.find("span", class_ = "gb-summary-grade-score-outof")
                        noWeightMarkers = assessments.find("span", class_ = "gb-flag-not-counted")

                        #parse elements in gradebook for desired data and perform calulations
                        #check if there is grade data, if not, assign no maximum grade
                        if assnGrade != "": 
                            assnMaxGrade = re.sub(r"/", "", assnMaxGrade.get_text())
                        else:
                            assnMaxGrade = ""
                        #check if a categoryWeight and "no weight flag" exists
                        if isinstance(categoryWeight[numberOfCats-1], float) == True and noWeightMarkers == None and assnGrade != "": 
                            assnWeight = categoryWeight[numberOfCats-1]/(len(allAssessments)-1-noGradeData) #calculate weight of individual assessment based on category weight
                            weightedGrade = (float(assnGrade)/float(assnMaxGrade)) * assnWeight
                        else: 
                            assnWeight = 0
                            weightedGrade = 0
                        
                        #append data to arrays
                        names.append(assessment + " ") #add space after name to differentiate between categories and assessments
                        marks.append(assnGrade) 
                        maxMark.append(assnMaxGrade)
                        weight.append(assnWeight)
                        weightedGradeCalc.append(weightedGrade)
            
            #for gradebooks with no assignment categories (html elements are differently named)
            else:
                
                #find assessments in no grade data to properly redistribute grade weights
                for noGradeAssessments in gbtable.findAll("tr", class_ = re.compile(r"^\bgb-summary-grade-row gb-no-categories (odd|even)\b$")):
                    noWeightMarkers = noGradeAssessments.find("span", class_ = "gb-flag-not-counted")
                    assnGrade = noGradeAssessments.find("span", class_ = "gb-summary-grade-score-raw").get_text()
                    if assnGrade == "" or noWeightMarkers != None:
                        noGradeData = noGradeData + 1
                
                #find individual assessments in gradebook
                for assessments in gbtable.findAll("tr", class_ = re.compile(r"^\bgb-summary-grade-row gb-no-categories (odd|even)\b$")):
                    #search for desired elements
                    assessment = assessments.find("span", class_ = "gb-summary-grade-title").get_text()
                    assnGrade = assessments.find("span", class_ = "gb-summary-grade-score-raw").get_text()
                    assnGradeAll = assessments.findAll("span", class_ = "gb-summary-grade-score-raw")
                    assnMaxGrade = assessments.find("span", class_ = "gb-summary-grade-score-outof")
                    noWeightMarkers = assessments.find("span", class_ = "gb-flag-not-counted")

                    #parse elements for desired data and perform calulations
                     #check if there is grade data, otherwise, assign "" to maximum grade
                    if assnGrade != "":
                        assnMaxGrade = re.sub(r"/", "", assnMaxGrade.get_text())
                    else:
                        assnMaxGrade = ""
                    #check if a no weight flag exists
                    if noWeightMarkers == None and assnGrade != "": 
                        #ask user to input assignment weight if it is not given
                        assnWeight = input("Assessment weight for " + assessment + " in " + courseNames[i] + " was not provided. Please enter it now (in percent)\n")
                        while True:
                            if assnWeight.isdigit() == False:
                                print("Invalid input, please try again")
                                assnWeight = input()
                            elif float(assnWeight) < 0 or float(assnWeight) > 100:
                                print("Inputs must be between 0% - 100%, please try again")
                                assnWeight = input()
                            else:
                                assnWeight = float(assnWeight) * 0.01
                                break
                        weightedGrade = (float(assnGrade)/float(assnMaxGrade)) * assnWeight
                    else: 
                        assnWeight = 0
                        weightedGrade = 0
                    
                    #append data to arrays
                    names.append(assessment + " ") #add space after name to differentiate between categories and assessments
                    marks.append(assnGrade) 
                    maxMark.append(assnMaxGrade)
                    weight.append(assnWeight)
                    weightedGradeCalc.append(weightedGrade)
                    categoryWeight.append(assnWeight)

        #calculate and add the final course grade to arrays
        #check to see if course weights sum to 100.
        #if they don't warn the user and calculate grade based on total weight
        sumOfWeights = arrSum(categoryWeight)
        if sumOfWeights == 1:
            final = round(arrSum(weightedGradeCalc*100), 2)
            marks.append(final)
            finalGrades.append(final)
        elif sumOfWeights <= 0:
            print("Warning: no gradebook exists for " + courseNames[i] + ", grades could not be extracted")
            marks.append("")
            finalGrades.append("")
        else:
            print("Warning: the assignment weights for " + courseNames[i] +  " do not sum to 100. Course grade may be incorrect")
            final = round(arrSum(weightedGradeCalc*100)/arrSum(categoryWeight), 2)
            marks.append(final)
            finalGrades.append(final)
        
        #append headings for final course grade to arrays
        names.append("COURSE GRADE")
        maxMark.append("/100")
        weight.append("")
        weightedGradeCalc.append("")

        #round the grade weight and weighted grade to 2 decimal places, convert fraction to percent
        weight = roundArr(weight, 4)
        weightedGradeCalc = roundArr(weightedGradeCalc, 4)

        #set up the excel dataframe
        df = pd.DataFrame({'Gradebook Item':names,'Grade':marks,'Max Points':maxMark, 'Weight (out of 100%)':weight, 'Weighted Grade (out of 100%)':weightedGradeCalc})

        #define default style
        default_style = Styler(font_size = 12)
        #create a new StyleFrame with excel dataframe using default style
        sf[i] = StyleFrame(df, styler_obj = default_style)

        #apply header style
        header_style = Styler(bg_color = utils.colors.black, font_color = utils.colors.white, bold = True, font_size = 18, shrink_to_fit = False)
        sf[i].apply_headers_style(styler_obj = header_style)

        #apply style for categories. There is a different style for the headings (they are justified to the left)
        categoryStyleName = Styler(bold = True, font_size = 14, bg_color = utils.colors.grey, horizontal_alignment = utils.horizontal_alignments.left)
        categoryStyle = Styler(bold = True, font_size = 14, bg_color = utils.colors.grey)
        #prevents non-category items from receiving category formatting
        for j in range(0, len(categoryNames)):
            sf[i].apply_style_by_indexes(indexes_to_style = sf[i][sf[i]['Gradebook Item'] == categoryNames[j]], 
            overwrite_default_style = False, 
            styler_obj = categoryStyleName)
            sf[i].apply_style_by_indexes(indexes_to_style = sf[i][sf[i]['Gradebook Item'] == categoryNames[j]],
            cols_to_style = ('Grade', 'Max Points', 'Weight (out of 100%)', 'Weighted Grade (out of 100%)'),
            overwrite_default_style = False,
            styler_obj = categoryStyle)

        #apply style for final grade cell
        finalGradeStyle = Styler(bg_color = utils.colors.grey, bold = True, font_size = 16, horizontal_alignment = utils.horizontal_alignments.left)
        sf[i].apply_style_by_indexes(indexes_to_style = sf[i][sf[i]['Gradebook Item'] == 'COURSE GRADE'], 
        overwrite_default_style = False, 
        styler_obj = finalGradeStyle)

        #factors for auto-setting cell width
        StyleFrame.A_FACTOR = 7
        StyleFrame.P_FACTOR = 1.3

    else:
        finalGrades.append("")
        print("Warning: no gradebook exists for " + courseNames[i] + ", grades could not be extracted")

#setup dataframes to be written to excel sheet
writer = StyleFrame.ExcelWriter('Gradebook Rip.xlsx')
courseGradeCount = 0

#determine which courses have no grade data (used for GPA calculation)
for i in range(len(finalGrades)):
    if finalGrades[i] != "": 
        courseGradeCount = courseGradeCount + 1

#calculate GPA (out of 100%), assumes all courses have equal weight
if courseGradeCount > 0:
    GPA = round(arrSum(finalGrades)/courseGradeCount, 5)
else:
    GPA = ""
    print("Warning: no gradebook pages were found or there are no courses in favorites bar. GPA could not be calculated")
courseNames.append("GPA")
finalGrades.append(GPA)

#dataframe for GPA (overview) sheet
overviewDf = pd.DataFrame({"Course": courseNames, "Grade (out of 100%)": finalGrades})

#define default style
default_style = Styler(font_size = 12)
#create a new StyleFrame with excel dataframe using default style
overviewSf = StyleFrame(overviewDf, styler_obj = default_style)

#apply header style
header_style = Styler(bg_color = utils.colors.black, font_color = utils.colors.white, bold = True, font_size = 18, shrink_to_fit = False)
overviewSf.apply_headers_style(styler_obj = header_style)

#apply style for GPA cell
finalGradeStyle = Styler(bg_color = utils.colors.grey, bold = True, font_size = 16, horizontal_alignment = utils.horizontal_alignments.left)
overviewSf.apply_style_by_indexes(indexes_to_style = overviewSf[overviewSf['Course'] == 'GPA'], 
overwrite_default_style = False, 
styler_obj = finalGradeStyle)

#write GPA and course grades to sheet
overviewSf.to_excel(writer, sheet_name = "Overview", best_fit = ("Course", "Grade (out of 100%)"))

#write course data to excel sheet
for i in range(len(sf)):
    if sf[i] != None:
        sf[i].to_excel(writer, sheet_name = courseNames[i], best_fit = ('Gradebook Item', 'Grade', 'Max Points', 'Weight (out of 100%)', 'Weighted Grade (out of 100%)'))

#save and close excel sheet
writer.save()
writer.close()
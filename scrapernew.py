import requests
import urllib.request
import time
import xlwt
from bs4 import BeautifulSoup

def addHeadersToSheet(worksheet):
    #Add Style for the Headers
    style_text_wrap_font_bold_black_color = xlwt.easyxf('align:wrap on; font: bold on, color-index black')
    col_width = 128*30
    worksheet.write(0, 0, "BREED", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 1, "HEIGHT", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 2, "WEIGHT", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 3, "LIFE EXPECTANCY", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 4, "CHARACTERISTICS", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 5, "GROOMING FREQUENCY", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 6, "SHEDDING LEVEL", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 7, "ENERGY LEVEL", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 8, "TRAINABILITY", style_text_wrap_font_bold_black_color)
    worksheet.write(0, 9, "TEMPERAMENT/DEMEANOR", style_text_wrap_font_bold_black_color)

    worksheet.col(0).width = col_width
    worksheet.col(1).width = col_width
    worksheet.col(2).width = col_width
    worksheet.col(3).width = col_width
    worksheet.col(4).width = col_width
    worksheet.col(5).width = col_width
    worksheet.col(6).width = col_width
    worksheet.col(7).width = col_width
    worksheet.col(8).width = col_width
    worksheet.col(9).width = col_width

def insertDataInSheet(worksheet, currentDogCounter, dog):
    breed = dog.find("div", {"id": "page-title"}).select('h1')[0].text.strip()

    print(str(currentDogCounter) + " " + breed)

    listItems = ["NA", "NA", "NA", "NA"]
    attributeList = dog.find("ul", {"class": "attribute-list"})
    for li in attributeList.find_all("li"):
        data = li.find_all("span")
        if (data[0].string == "Temperament:"):
            listItems[0] = data[1].string
        elif (data[0].string == "Height:"):
            listItems[1] = data[1].string
        elif (data[0].string == "Weight:"):
            listItems[2] = data[1].string
        elif (data[0].string == "Life Expectancy:"):
            listItems[3] = data[1].string

    groomingFrequency = "NA"
    shedding = "NA"
    try:
        groomingTab = dog.find("div", {"id": "panel-GROOMING"}).find_all("div", {"class": "graph-section__inner"})
        for option in groomingTab:
            if (option.find("h4").string == "Grooming Frequency"):
                groomingFrequency = option.find("div", {"class": "bar-graph__text"}).string
            elif (option.find("h4").string == "Shedding"):
                shedding = option.find("div", {"class": "bar-graph__text"}).string
    except AttributeError:
        x = "NA"

    energyTab = dog.find("div", {"id": "panel-EXERCISE"})

    try:
        energyLevel = energyTab.find_all("div", {"class": "graph-section__inner"})[0].find("div", {"class": "bar-graph__text"}).string
    except IndexError:
        energyLevel = "DOUBLE CHECK"
    except AttributeError:
        energyLevel = "NA"

    trainability = "NA"
    temperament = "NA"
    try:
        trainingTab = dog.find("div", {"id": "panel-TRAINING"}).find_all("div", {"class": "graph-section__inner"})
        for option in trainingTab:
            if (option.find("h4").string == "Trainability"):
                trainability = option.find("div", {"class": "bar-graph__text"}).string
            elif (option.find("h4").string == "Temperament/Demeanor"):
                temperament = option.find("div", {"class": "bar-graph__text"}).string
            else:
                print("none found for training")
    except AttributeError:
        x = "NA"

    characteristics = listItems[0]
    height = listItems[1]
    weight = listItems[2]
    lifeExpancy = listItems[3]

    worksheet.write(currentDogCounter, 0, breed)
    worksheet.write(currentDogCounter, 1, listItems[1])
    worksheet.write(currentDogCounter, 2, listItems[2])
    worksheet.write(currentDogCounter, 3, listItems[3])
    worksheet.write(currentDogCounter, 4, listItems[0])
    worksheet.write(currentDogCounter, 5, groomingFrequency)
    worksheet.write(currentDogCounter, 6, shedding)
    worksheet.write(currentDogCounter, 7, energyLevel)
    worksheet.write(currentDogCounter, 8, trainability)
    worksheet.write(currentDogCounter, 9, temperament)

#Set Up the Excel File
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet("Dogs")
excel_file_path = "./Dog Options.xls"

addHeadersToSheet(worksheet)

currentDogCounter = 1
for i in range(24):
    url = "https://www.akc.org/dog-breeds/page/" + str(i + 1)
    response = requests.get(url)

    soup = BeautifulSoup(response.text, "lxml")

    topDiv = soup.find("div", {"class": "contents-grid-group"})
    secondDiv = topDiv.find("div")
    dogChoices = secondDiv.find_all("div", {"class": "grid-col"})

    for dog in dogChoices:
        href = dog.find("a").get("href")
        nextResponse = requests.get(href)

        dog = BeautifulSoup(nextResponse.text, "lxml")
        insertDataInSheet(worksheet, currentDogCounter, dog)
        currentDogCounter += 1

workbook.save(excel_file_path)

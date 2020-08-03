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

    attributeList = dog.find("ul", {"class": "attribute-list"})

    try:
        characteristics = attributeList.find_all("li")[0].find("span", {"class": "attribute-list__description"}).string
    except IndexError:
        characteristics = "NA"
    except AttributeError:
        characteristics = "NA"
    try:
        height = attributeList.find_all("li")[2].find("span", {"class": "attribute-list__description"}).string
    except IndexError:
        height = "NA"
    except AttributeError:
        height = "NA"
    try:
        weight = attributeList.find_all("li")[3].find("span", {"class": "attribute-list__description"}).string
    except IndexError:
        weight = "NA"
    except AttributeError:
        weight = "NA"
    try:
        lifeExpancy = attributeList.find_all("li")[4].find("span", {"class": "attribute-list__description"}).string
    except IndexError:
        lifeExpancy = "NA"
    except AttributeError:
        lifeExpancy = "NA"

    groomingTab = dog.find("div", {"id": "panel-GROOMING"})

    try:
        groomingFrequency = groomingTab.find_all("div", {"class": "graph-section__inner"})[0].find("div", {"class": "bar-graph__text"}).string
    except IndexError:
        groomingFrequency = "NA"
    except AttributeError:
        groomingFrequency = "NA"
    try:
        shedding = groomingTab.find_all("div", {"class": "graph-section__inner"})[1].find("div", {"class": "bar-graph__text"}).string
    except IndexError:
        shedding = "NA"
    except AttributeError:
        shedding = "NA"

    energyTab = dog.find("div", {"id": "panel-EXERCISE"})

    try:
        energyLevel = energyTab.find_all("div", {"class": "graph-section__inner"})[0].find("div", {"class": "bar-graph__text"}).string
    except IndexError:
        energyLevel = "DOUBLE CHECK"
    except AttributeError:
        energyLevel = "NA"

    trainingTab = dog.find("div", {"id": "panel-TRAINING"})

    try:
        trainability = trainingTab.find_all("div", {"class": "graph-section__inner"})[0].find("div", {"class": "bar-graph__text"}).string
    except IndexError:
        trainability = "DOUBLE CHECK"
    except AttributeError:
        trainability = "NA"
    try:
        temperament = trainingTab.find_all("div", {"class": "graph-section__inner"})[1].find("div", {"class": "bar-graph__text"}).string
    except IndexError:
        temperament = "DOUBLE CHECK"
    except AttributeError:
        temperament = "NA"

    worksheet.write(currentDogCounter, 0, breed)
    worksheet.write(currentDogCounter, 1, height)
    worksheet.write(currentDogCounter, 2, weight)
    worksheet.write(currentDogCounter, 3, lifeExpancy)
    worksheet.write(currentDogCounter, 4, characteristics)
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

from itertools import count
from urllib.request import urlopen,Request
from bs4 import BeautifulSoup
import openpyxl as xl
from openpyxl.styles import Font

# scrape the website below to retrieve the top 5 countries with the highest GDPs. Calculate the GDP per capita
# by dividing the GDP by the population. You can perform the calculation in Python natively or insert the code
# in excel that will perform the calculation in Excel by each row. DO NOT scrape the GDP per capita from the
# webpage, make sure you use your own calculation.

# FOR YOUR REFERENCE - https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/numbers.html
# this link shows you the different number formats you can apply to a column using openpyxl


# FOR YOUR REFERENCE - https://www.geeksforgeeks.org/python-string-replace/
# this link shows you how to use the REPLACE function (you may need it if your code matches mine but not required)

### REMEMBER ##### - your output should match the excel file (GDP_Report.xlsx) including all formatting.

webpage = 'https://www.worldometers.info/gdp/gdp-by-country/'

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}
req = Request(webpage, headers=headers)
webpage = urlopen(req).read()			

soup = BeautifulSoup(webpage, 'html.parser')

title = soup.title

#print(title.text)

table_rows = soup.findAll("tr")

#empty lists to store all column data bc doing rank = td[0] says list index out of range
rank = []
country = []
country_GDP = []
country_GDP_abbrv = []
GDP_growth = []
population = []
GDP_per_capita_given = []
share_of_world_GDP = []


for row in table_rows[0:6]:
    td = row.findAll("td")
    counter = 0
    for t in td:
        t = t.text.strip()

        if counter == 0:
            t = int(t)
            rank.append(t)

        if counter == 1:
            country.append(t)
        
        if counter == 2:
            t = t.replace(",","")
            t = t.replace("$","")
            t = int(t)
            country_GDP.append(t)

        if counter == 3:
            country_GDP_abbrv.append(t)
        
        if counter == 4:
            GDP_growth.append(t)
        
        if counter == 5:
            t = t.replace(",","")
            t = int(t)
            population.append(t)
        
        if counter == 6:
            GDP_per_capita_given.append(t)
        
        if counter == 7:
            share_of_world_GDP.append(t)

        counter += 1

#create a new excel document
wb = xl.Workbook()

MySheet = wb.active

MySheet.title = "First Sheet"

#write content to a cell
MySheet['A1']='No.'
MySheet['B1']='Country'
MySheet['C1']='GDP'
MySheet['D1']='Population'
MySheet['E1']='GDP Per Capita'

#change the font size and italicize
MySheet['A1'].font = Font(name="Calibri", size=16, italic=False,bold=True)
MySheet['B1'].font = Font(name="Calibri", size=16, italic=False,bold=True)
MySheet['C1'].font = Font(name="Calibri", size=16, italic=False,bold=True)
MySheet['D1'].font = Font(name="Calibri", size=16, italic=False,bold=True)
MySheet['E1'].font = Font(name="Calibri", size=16, italic=False,bold=True)

#adding values to cells
MySheet['A2'] = rank[0]
MySheet['A3'] = rank[1]
MySheet['A4'] = rank[2]
MySheet['A5'] = rank[3]
MySheet['A6'] = rank[4]

MySheet['B2'] = country[0]
MySheet['B3'] = country[1]
MySheet['B4'] = country[2]
MySheet['B5'] = country[3]
MySheet['B6'] = country[4]

MySheet['C2'] = country_GDP[0]
MySheet['C3'] = country_GDP[1]
MySheet['C4'] = country_GDP[2]
MySheet['C5'] = country_GDP[3]
MySheet['C6'] = country_GDP[4]

MySheet['D2'] = population[0]
MySheet['D3'] = population[1]
MySheet['D4'] = population[2]
MySheet['D5'] = population[3]
MySheet['D6'] = population[4]

MySheet['E2'] = '=(C2/D2)'
MySheet['E3'] = '=(C3/D3)'
MySheet['E4'] = '=(C4/D4)'
MySheet['E5'] = '=(C5/D5)'
MySheet['E6'] = '=(C6/D6)'


#change the column width
MySheet.column_dimensions['A'].width = 1
MySheet.column_dimensions['B'].width = 14
MySheet.column_dimensions['A'].width = 25
MySheet.column_dimensions['A'].width = 20
MySheet.column_dimensions['A'].width = 25



wb.save("PythonToExcel.xlsx")
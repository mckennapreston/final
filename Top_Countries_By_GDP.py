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

req = Request(webpage, headers = headers)

#page = urlopen(webpage)	
webpage = urlopen(req).read()

soup = BeautifulSoup(webpage, 'html.parser')

title = soup.title

GDP_table = soup.find('table')
GDP_rows = GDP_table.findAll('tr')
#print(GDP_rows[1])


#############################

wb = xl.Workbook()
MySheet = wb.active

MySheet.title = 'GDP Report'

MySheet['A1'] = 'No.'
MySheet['B1'] = 'Country'
MySheet['C1'] = 'GDP'
MySheet['D1'] = 'Population'
MySheet['E1'] = 'GDP Per Capita'

for x in range(1,6):
    td = GDP_rows[x].findAll('td')
    ranking = td[0].text
    country = td[1].text
    GDPnum = td[2].text.replace(",","").replace("$","")
    population = int(float(td[5].text.replace(",","")))
    

    GDPperCapita = int(GDPnum)/population
    GDP_perCapita = "${:,.2f}".format(GDPperCapita)
    #print('$' + format(GDPperCapita, ',.2f'))

    MySheet['A' + str(x+1)] = ranking
    MySheet['B' + str(x+1)] = country
    MySheet['C' + str(x+1)] = '$' + GDPnum
    MySheet['D' + str(x+1)] = population
    MySheet['E' + str(x+1)] = GDPperCapita
    

#headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}
#req = Request(webpage, headers=headers)
#webpage = urlopen(req).read()			

MySheet.column_dimensions['A'].width = 5
MySheet.column_dimensions['B'].width = 15
MySheet.column_dimensions['C'].width = 25
MySheet.column_dimensions['D'].width = 16
MySheet.column_dimensions['E'].width = 20


header_font = Font(size=16, bold = True)

for cell in MySheet[1:1]:
    cell.font = header_font

wb.save('GDPReport2.xlsx')


    






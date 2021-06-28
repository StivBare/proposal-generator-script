import xlwings as xw
import json
import math


DATA_PATH = ""
TEMPLATE_PATH = ""

# ROUND UP VALUE TO THE NEAREST 10
def roundup(x):
    return int(math.ceil(x/10) * 10)

# EXTRACT DATA FROM EXPORTED JSON FILE
f = open(DATA_PATH)
data = json.load(f)
data_roof = data["roof"]

area = int(math.ceil(data_roof["roof_facets"]["area"] / 100))
perimeter = roundup(data_roof["rakes"]["length"] + data_roof["gutters_eaves"]["length"])
eaves = roundup(data_roof["gutters_eaves"]["length"])
valleys = roundup(data_roof["valleys"]["length"])


# WRITE DATA TO TEMPLATE
workbook = xw.Book(TEMPLATE_PATH)
data_entry = workbook.sheets["Data_Entry"]

data_entry.range("F22").value = area
data_entry.range("F25").value = perimeter
data_entry.range("F26").value = eaves
data_entry.range("F27").value = valleys

# workbook.save(TEMPLATE_PATH)





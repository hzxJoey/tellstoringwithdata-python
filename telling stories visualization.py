import numpy as np
import xlrd as xl
from pyecharts.charts import Bar
from pyecharts import options as opts

# import the xlsx file
workSheet = xl.open_workbook("/Users/hezixi/Downloads/CBPSN04304.download.xlsx")
# import the sheets' name
name = workSheet.sheet_names()
# get the sheet we want to use
sexWorkSheet = workSheet.sheet_by_index(3)
nRows = sexWorkSheet.nrows
nCols = sexWorkSheet.ncols
# convert worksheet to array
tableArray = np.array([sexWorkSheet.row_values(0)])
for i in range(1, nRows):
    tRows = np.array([sexWorkSheet.row_values(i)])
    tableArray = np.concatenate((tableArray, tRows), axis=0)
# convert male datatype to float
tempM = tableArray[4:17, 1:21]
tempM[tempM == '-'] = 0
maleData = tempM.astype(float)

# convert female datatype to float
tempW1 = tableArray[20:32, 1:21]
tempW2 = [tableArray[33, 1:21]]
tempW = np.concatenate((tempW1, tempW2), axis=0)
tempW[tempW == '-'] = 0
femaleData = tempW.astype(float)

# convert all datatype to float
tempA = tableArray[37:50, 1:21]
tempA[tempA == '-'] = 0
allData = tempA.astype(float)

# plot the data
labelListYear = list(tableArray[2, 13:21])
maleList = list(maleData[12, 12:])
femaleList = list(femaleData[12, 12:])
allList = list(allData[12,12:])
bar = (
    Bar()
    .add_xaxis(labelListYear)
    .add_yaxis('maleVictims',maleList)
    .add_yaxis('femaleVctims',femaleList)
    .add_yaxis('allVictims',allList)
    .set_global_opts(title_opts=opts.TitleOpts(title='Victims distinguished by gender'))
)

bar.render()
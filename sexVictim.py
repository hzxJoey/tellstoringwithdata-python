import numpy as np
import matplotlib.pyplot as plt
import xlrd as xl

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
plt.figure(figsize=(15,10))
labelListYear = list(tableArray[2, 1:21])
maleList = list(maleData[12, :])
femaleList = list(femaleData[12, :])
x = np.arange(len(labelListYear))
barw = 0.5
plt.barh(x, maleList, height=0.4, color='purple', label='maleVictim')
plt.barh(x + barw, femaleList, height=0.4, color='turquoise', label='femaleVictim')
for a, b in zip(x, maleList):
    plt.text(b, a+0.05, '%.2f' % b, ha='center', va='bottom', fontsize=12)
for c, d in zip(x+barw, femaleList):
    plt.text(d, c+0.05, '%.2f' % d, ha='center', va='bottom', fontsize=12)
plt.legend()
plt.yticks(x + barw / 2, labelListYear)
plt.xlabel('amount')
plt.ylabel('Year')
plt.title('victims distinguished by gender')
plt.show()

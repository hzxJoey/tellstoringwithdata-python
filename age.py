import numpy as np
import matplotlib.pyplot as plt
import xlrd as xl

# import the xlsx file
workSheet = xl.open_workbook("/Users/hezixi/Downloads/CBPSN04304.download.xlsx")
# import the sheets' name
name = workSheet.sheet_names()
print name
# get the sheet we want to use
ageWorkSheet = workSheet.sheet_by_index(7)
nRows = ageWorkSheet.nrows
nCols = ageWorkSheet.ncols
# convert worksheet to array
tableArray = np.array([ageWorkSheet.row_values(0)])
for i in range(1, nRows):
    tRows = np.array([ageWorkSheet.row_values(i)])
    tableArray = np.concatenate((tableArray, tRows), axis=0)
# select the age group
age_arrayT = tableArray[3:49, 2:4]
age_array = age_arrayT.astype(float)
timeList = list(tableArray[3:49, 0])
# create plot
n = np.arange(len(timeList))
plt.figure(figsize=(35, 10))
plt.ylim(650, 6000)
plt.plot(n, age_array.T[0], 'c--', linewidth=5, label='age 10-17')
plt.plot(n, age_array.T[1], 'k--', linewidth=5, label='age 18+')
for a, b in zip(n, age_array.T[0]):
    plt.text(a, b, b, ha='center', va='bottom', fontsize=13)
for a, b in zip(n, age_array.T[1]):
    plt.text(a, b, b, ha='center', va='bottom', fontsize=13)
plt.xticks(n, timeList)
plt.xlabel('Times')
plt.ylabel('amount')
plt.title(' OFFENCES INVOLVING THE POSSESSION OF A KNIFE OR OFFENSIVE WEAPON')
plt.legend()
plt.show()


# An experiment python program
import pdb
import numpy as np
import pandas as pd

print ("Let's process an Excel sheet !")
print ("The Panda version is: ", pd.__version__)

a = np.array([1, 2, 3, 4, 5])
b = np.argwhere(a > 6)
c = 0
if b.size != 0:
    c = 1

print(a)
print(c)
#c = a[b]
#print(c)

#TargetSheet = pd.read_excel("sample.xlsx")

# head() is the first row of panda DataFrame:
# An parameter n > 0 it will give the first n rows
# An parameter n < 0 it will give the first (Total rows + n) rows
#print(TargetSheet.head(3))   

#print(TargetSheet.columns)
#print(TargetSheet.index)
#print(TargetSheet.loc[4, 'Unnamed: 3'])   # Get value based on the row/column index object

# Get value based on the row/colum position
# index = 0 starts with the second row/column in excel sheet. 
# Panda uses the first row/column to lable each cell.

#print(TargetSheet.iloc[4, 3])  
#print(TargetSheet.iloc[4, :])  # Slice opearations
#print(TargetSheet)


#pdb.set_trace()



#Dimension = TargetSheet.shape
#print ("Dimension is", Dimension)



# Write to the output Excel sheet

#WriteObj = pd.DataFrame([['a', 'b'], ['c', 'd']], 
#                        index=['row 1', 'row 2'], 
#                        columns=['col 1', 'col 2'])

#WriteObj.to_excel("output.xlsx", sheet_name='sheet1')


#pdb.set_trace()








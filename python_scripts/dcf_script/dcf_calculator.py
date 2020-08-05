# This is a DCF calculator for getting a price range of the target stock
# Date: 06/16/2020

import pdb
import numpy as np
import pandas as pd

#pdb.set_trace()
print ("******** DCF calculator started ********")
print ("The Panda version used is: ", pd.__version__)

# Read in source sheet
TargetSheet = pd.read_excel("sample.xlsx")

# Print shape of the sheet
Dimension = TargetSheet.shape
print ("Dimension is", Dimension)

# Get the revenue, net income and free cash flow vector and other parameters from source
ReqRate = TargetSheet.iloc[0, 1]    # for a scalar value no need to call .values() method to get its cell value

ShareVol = TargetSheet.iloc[1, 1]

PerpGrowRate = TargetSheet.iloc[2, 1]

CurrentRevObj = TargetSheet.iloc[6, 1:5]  # slice end index need to +1
CurrentRev = CurrentRevObj.values
#print ("CurrentRev = ", CurrentRev) 

CurrentNetObj = TargetSheet.iloc[7, 1:5]
CurrentNet = CurrentNetObj.values

CurrentCashObj = TargetSheet.iloc[8, 1:5]
CurrentCash = CurrentCashObj.values

FutureRevObj = TargetSheet.iloc[11, 1:3]
FutureRev = FutureRevObj.values


# Calculate Net Profit Margin (Npm) and Free Cash Flow to Margin (Fcfm)
Npm = CurrentNet/CurrentRev
Fcfm = CurrentCash/CurrentNet
ConfidFactor = 1.4   # Set Confidential Interval to 1.7 times Std
EvalFactor = 0.9     # Set a conservative factor of 0.9

# Post processing of noisy data
NpmRawMean = Npm.mean()
NpmStd = Npm.std()
NpmRangeMax = NpmRawMean + ConfidFactor * NpmStd
NpmRangeMin = NpmRawMean - ConfidFactor * NpmStd

#NpmPost = Npm.clip(min=NpmRangeMin, max=NpmRangeMax)
NpmDelete = np.argwhere(((Npm > NpmRangeMax) | (Npm < NpmRangeMin)) & (Npm > 0))

# Log the abnormal data of Net Profit
if NpmDelete.size != 0:
    print("Log:: The following Net Profit Margin is abnormal and will be excluded:")
    for irow in range(0, NpmDelete.size):
        DeleIdx = NpmDelete[irow, 0]
        YearIdx = DeleIdx + 1
        print("Log:: Year", YearIdx, "Rate =", Npm[DeleIdx],"is excluded")

NpmPost = np.delete(Npm, NpmDelete)
NpmMean = NpmPost.mean()
NpmActual = NpmMean * EvalFactor


FcfmRawMean = Fcfm.mean()
FcfmStd = Fcfm.std()
FcfmRangeMax = FcfmRawMean + ConfidFactor * FcfmStd
FcfmRangeMin = FcfmRawMean - ConfidFactor * FcfmStd

FcfmDelete = np.argwhere(((Fcfm > FcfmRangeMax) | (Fcfm < FcfmRangeMin)) & (Fcfm > 0) )

# Log the abnormal data of FCF
if FcfmDelete.size != 0:
    print("Log:: The following FCF/Margin rate is abnormal and will be excluded:")
    for irow in range(0, FcfmDelete.size):
        DeleIdx = FcfmDelete[irow, 0]
        YearIdx = DeleIdx + 1
        print("Log:: Year", YearIdx, "Rate =", Fcfm[DeleIdx],"is excluded")

FcfmPost = np.delete(Fcfm, FcfmDelete)
FcfmMean = FcfmPost.mean()
FcfmActual = FcfmMean * EvalFactor


# Derive the Discount Factor (Df)
Df = np.ndarray(shape=(5,), dtype=float)

for i in range(0, 4):
    Df_elem = (1 + ReqRate)**(i+1)
    Df[i] = Df_elem

Df[4] = Df[3]  # Repeat the last elem
#print ("Df = ", Df)

# DCF calculations, loop over different EstGrowRate

EstGrowRateArr = np.array([0.5, 0.3, 0.2, 0.15, 0.1, 0.05, 0.02])
PricePerShareArr = np.empty(shape=(7,))
arr_indx = 0

for EstGrowRate in EstGrowRateArr:
    
    # Predictions of FCF and Net Income with furture revenues estimates

    # Extropolate future revenue
    FutureRevExp = FutureRev
    LatestEst = FutureRev[len(FutureRev) - 1]
    FirstPred = LatestEst * (1+EstGrowRate)
    SecPred = FirstPred * (1+EstGrowRate)
    FutureRevExp = np.append(FutureRevExp, FirstPred)
    FutureRevExp = np.append(FutureRevExp, SecPred)
    #print("FirstPred = ", FirstPred)
    #print("SecPred = ", SecPred)

    # Get Net Income estimate based on future revenue
    NetEst = FutureRevExp * NpmActual
    FcfPred = NetEst * FcfmActual

    # Terminal Value of FCF
    FcfTerm = FcfPred[len(FcfPred)-1] * (1 + PerpGrowRate)/(ReqRate - PerpGrowRate)
    FcfPred = np.append(FcfPred, FcfTerm)

    # Calcualte PV of FCF
    FcfPv = FcfPred/Df
    FcfSum = FcfPv.sum()
    #print("FcfSum = ", FcfSum)

    # Calculate Share Price
    PricePerShare = FcfSum/ShareVol
    PricePerShareArr[arr_indx] = round(PricePerShare, 2)

    arr_indx = arr_indx + 1

#print("PricePerShareArr = ", PricePerShareArr)

# Post processing and write to output xlsx sheet

EstGrowRateArrList = EstGrowRateArr.tolist()
GrowRateTable = []
RawRateNpmTable = [format(NpmRawMean, ".0%"), 0, 0, 0, 0, 0, 0]
AcutualRateNpmTable = [format(NpmActual, ".0%"), 0, 0, 0, 0, 0, 0]
RawRateFcfmTable = [format(FcfmRawMean, ".0%"), 0, 0, 0, 0, 0, 0]
ActualFcfmRateTable = [format(FcfmActual, ".0%"), 0, 0, 0, 0, 0, 0]

for elem in EstGrowRateArrList:
    GrowRateTable.append(format(elem, ".0%"))

WriteObj = pd.DataFrame([PricePerShareArr, RawRateNpmTable, AcutualRateNpmTable, RawRateFcfmTable, ActualFcfmRateTable], 
                        index=['Share Price', 'Raw Profit Margin', 'Actual Proft Margin', 'Raw FCF/Margin', 'Actual FCF/Margin'], 
                        columns=GrowRateTable)

WriteObj.to_excel("output.xlsx", sheet_name='sheet1')

print("******** END: output.xlsx has been generated ********")
# %%
from Imports import *

fmData = ""
em6PData = ""

totalAverage1 = []
totalAverage2 = []
totalAverage3 = []
totalAverage4 = []
totalAverage5 = []
totalAverage6 = []
totalAverage7 = []

em6PTotalAverage3 = []
em6PTotalAverage4 = []
em6PTotalAverage5 = []

em6PDataAverage = []
em6PDataTotalAverage = []

inspectionData = ""

def ReadFM():
    global fmData

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    vt1Directory = (fr'\\192.168.2.19\quality control\2024\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\CRONICS\~DATA TREND 2024')
    os.chdir(vt1Directory)

    wb = load_workbook(filename='FM05000102 NEW.xlsm', data_only=True)
    sheet = wb['format']
    fmData = pd.DataFrame(sheet.values)
    fmData = fmData.replace(r'\s+', '', regex=True)

def ReadEM6P():
    global em6PData

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    vt1Directory = (fr'\\192.168.2.19\quality control\2024\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\DHYE\~New Trend 2024')
    os.chdir(vt1Directory)

    wb = load_workbook(filename='EM0580106P NEW.xlsm', data_only=True)
    sheet = wb['format']
    em6PData = pd.DataFrame(sheet.values)
    em6PData = em6PData.replace(r'\s+', '', regex=True)

def GettingFM(lotNumber):
    global fmData
    
    global totalAverage1
    global totalAverage2
    global totalAverage3
    global totalAverage4
    global totalAverage5
    global totalAverage6
    global totalAverage7

    global inspectionData

    average1 = 0
    average2 = 0
    average3 = 0
    average4 = 0
    average5 = 0
    average6 = 0
    average7 = 0

    totalAverage1 = []
    totalAverage2 = []
    totalAverage3 = []
    totalAverage4 = []
    totalAverage5 = []
    totalAverage6 = []
    totalAverage7 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in fmData.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = fmData.iloc[max(0, hiblowRow[0] - 3):min(len(fmData), hiblowRow[0] + 10), fmData.columns.get_loc(hiblowColumn[0]):fmData.columns.get_loc(hiblowColumn[0]) + 999]
    hiblowFiltered

    #Getting The Row, Column Location Of Lot Number
    findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    lotNumberRow = [index for index, _ in findLotNumber]
    lotNumberColumn = [column for _, column in findLotNumber]

    print("Row indices:", lotNumberRow)
    print("Column names:", lotNumberColumn)

    for a in range(0, len(lotNumberColumn)):
        # Get The Neighboring Data of Lot Number
        inspectionData = fmData.iloc[max(0, lotNumberRow[a]):min(len(fmData), lotNumberRow[a] + 10), fmData.columns.get_loc(lotNumberColumn[a]):fmData.columns.get_loc(lotNumberColumn[a]) + 5]

        average1 = inspectionData.iloc[3].mean()
        average2 = inspectionData.iloc[4].mean()
        average3 = inspectionData.iloc[5].mean()
        average4 = inspectionData.iloc[6].mean()
        average5 = inspectionData.iloc[7].mean()
        average6 = inspectionData.iloc[8].mean()
        average7 = inspectionData.iloc[9].mean()

        totalAverage1.append(average1)
        totalAverage2.append(average2)
        totalAverage3.append(average3)
        totalAverage4.append(average4)
        totalAverage5.append(average5)
        totalAverage6.append(average6)
        totalAverage7.append(average7)

        inspectionData

    totalAverage1 = statistics.mean(totalAverage1)
    totalAverage2 = statistics.mean(totalAverage2)
    totalAverage3 = statistics.mean(totalAverage3)
    totalAverage4 = statistics.mean(totalAverage4)
    totalAverage5 = statistics.mean(totalAverage5)
    totalAverage6 = statistics.mean(totalAverage6)
    totalAverage7 = statistics.mean(totalAverage7)

    totalAverage1 = f"{totalAverage1:.2f}"
    totalAverage2 = f"{totalAverage2:.2f}"
    totalAverage3 = f"{totalAverage3:.2f}"
    totalAverage4 = f"{totalAverage4:.2f}"
    totalAverage5 = f"{totalAverage5:.2f}"
    totalAverage6 = f"{totalAverage6:.2f}"
    totalAverage7 = f"{totalAverage7:.2f}"

    print(totalAverage1)
    print(totalAverage2)
    print(totalAverage3)
    print(totalAverage4)
    print(totalAverage5)
    print(totalAverage6)
    print(totalAverage7)

def GettingEM6P(lotNumber):
    global em6PData
    
    global em6PTotalAverage3
    global em6PTotalAverage4
    global em6PTotalAverage5

    global inspectionData

    average3 = 0
    average4 = 0
    average5 = 0

    em6PTotalAverage3 = []
    em6PTotalAverage4 = []
    em6PTotalAverage5 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in em6PData.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = em6PData.iloc[max(0, hiblowRow[0] - 3):min(len(em6PData), hiblowRow[0] + 10), em6PData.columns.get_loc(hiblowColumn[0]):em6PData.columns.get_loc(hiblowColumn[0]) + 999]
    hiblowFiltered

    #Getting The Row, Column Location Of Lot Number
    findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    lotNumberRow = [index for index, _ in findLotNumber]
    lotNumberColumn = [column for _, column in findLotNumber]

    print("Row indices:", lotNumberRow)
    print("Column names:", lotNumberColumn)

    inspectionData = em6PData.iloc[max(0, lotNumberRow[0]):min(len(em6PData), lotNumberRow[0] + 10), em6PData.columns.get_loc(lotNumberColumn[0]):em6PData.columns.get_loc(lotNumberColumn[0]) + 5]

    for a in range(0, len(lotNumberColumn)):
        # Get The Neighboring Data of Lot Number
        inspectionData = em6PData.iloc[max(0, lotNumberRow[a]):min(len(em6PData), lotNumberRow[a] + 10), em6PData.columns.get_loc(lotNumberColumn[a]):em6PData.columns.get_loc(lotNumberColumn[a]) + 5]

        average3 = inspectionData.iloc[5].mean()
        average4 = inspectionData.iloc[6].mean()
        average5 = inspectionData.iloc[7].mean()

        em6PTotalAverage3.append(average3)
        em6PTotalAverage4.append(average4)
        em6PTotalAverage5.append(average5)
    
    em6PTotalAverage3 = statistics.mean(em6PTotalAverage3)
    em6PTotalAverage4 = statistics.mean(em6PTotalAverage4)
    em6PTotalAverage5 = statistics.mean(em6PTotalAverage5)

    em6PTotalAverage3 = f"{em6PTotalAverage3:.2f}"
    em6PTotalAverage4 = f"{em6PTotalAverage4:.2f}"
    em6PTotalAverage5 = f"{em6PTotalAverage5:.2f}"

    print(em6PTotalAverage3)
    print(em6PTotalAverage4)
    print(em6PTotalAverage5)

# ReadFM()
# GettingFM("080224A-40")
# ReadEM6P()
# GettingEM6P("CAT-4A04DI")
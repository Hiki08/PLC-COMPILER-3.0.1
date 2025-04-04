# %%
from Imports import *

dfVt1 = ""

average1 = ""
average2 = ""
average3 = ""
average4 = ""
average5 = ""
average6 = ""
average7 = ""

inspectionData = ""
# %%
def ReadInspectionData():
    global dfVt1

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    vt1Directory = (fr'\\192.168.2.19\quality control\2024\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\CRONICS\~DATA TREND 2024')
    os.chdir(vt1Directory)

    wb = load_workbook(filename='FM05000102 NEW.xlsm', data_only=True)
    sheet = wb['format']
    dfVt1 = pd.DataFrame(sheet.values)
# %%
def GettingInspectionData(lotNumber):
    global dfVt1
    
    global average1
    global average2
    global average3
    global average4
    global average5
    global average6
    global average7

    global inspectionData

    average1 = 0
    average2 = 0
    average3 = 0
    average4 = 0
    average5 = 0
    average6 = 0
    average7 = 0

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in dfVt1.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = dfVt1.iloc[max(0, hiblowRow[0] - 3):min(len(dfVt1), hiblowRow[0] + 10), dfVt1.columns.get_loc(hiblowColumn[0]):dfVt1.columns.get_loc(hiblowColumn[0]) + 999]
    hiblowFiltered

    #Getting The Row, Column Location Of Lot Number
    findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    lotNumberRow = [index for index, _ in findLotNumber]
    lotNumberColumn = [column for _, column in findLotNumber]

    print("Row indices:", lotNumberRow)
    print("Column names:", lotNumberColumn)

    # Get The Neighboring Data of Lot Number
    inspectionData = dfVt1.iloc[max(0, lotNumberRow[0]):min(len(dfVt1), lotNumberRow[0] + 10), dfVt1.columns.get_loc(lotNumberColumn[0]):dfVt1.columns.get_loc(lotNumberColumn[0]) + 5]

    average1 = inspectionData.iloc[3].mean()
    average2 = inspectionData.iloc[4].mean()
    average3 = inspectionData.iloc[5].mean()
    average4 = inspectionData.iloc[6].mean()
    average5 = inspectionData.iloc[7].mean()
    average6 = inspectionData.iloc[8].mean()
    average7 = inspectionData.iloc[9].mean()

    average1 = f"{average1:.2f}"
    average2 = f"{average2:.2f}"
    average3 = f"{average3:.2f}"
    average4 = f"{average4:.2f}"
    average5 = f"{average5:.2f}"
    average6 = f"{average6:.2f}"
    average7 = f"{average7:.2f}"

    # print(average1)
    # print(average2)
    # print(average3)
    # print(average4)
    # print(average5)
    # print(average6)
    # print(average7)

    inspectionData
ReadInspectionData()
GettingInspectionData("072524A-40")

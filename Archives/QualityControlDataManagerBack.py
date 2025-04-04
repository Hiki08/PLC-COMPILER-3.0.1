# %%
from Imports import *

FM05000102Data = ""
EM0580106PData = ""
EM0580107PData = ""

totalAverage1 = []
totalAverage2 = []
totalAverage3 = []
totalAverage4 = []
totalAverage5 = []
totalAverage6 = []
totalAverage7 = []

inspectionData = ""

fmTotalMinimum1 = []
fmTotalMinimum2 = []
fmTotalMinimum3 = []
fmTotalMinimum4 = []
fmTotalMinimum5 = []
fmTotalMinimum6 = []
fmTotalMinimum7 = []

em6PTotalAverage3 = []
em6PTotalAverage4 = []
em6PTotalAverage5 = []
em6PTotalAverage10 = []

em6PTotalMinimum3 = []
em6PTotalMinimum4 = []
em6PTotalMinimum5 = []

em6PTotalMaximum3 = []
em6PTotalMaximum4 = []
em6PTotalMaximum5 = []

em7PTotalAverage3 = []
em7PTotalAverage4 = []
em7PTotalAverage5 = []
em7PTotalAverage10 = []

em7PTotalMinimum3 = []
em7PTotalMinimum4 = []
em7PTotalMinimum5 = []
em7PTotalMinimum10 = []


def ReadAll():
    ReadEM0580106P()
    ReadEM0580107P()
    ReadFM05000102()


def ReadEM0580106P():
    global EM0580106PData

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    vt1Directory = (fr'\\192.168.2.19\quality control\2024\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\DHYE\~New Trend 2024')
    os.chdir(vt1Directory)

    # wb = load_workbook(filename='EM0580106P NEW.xlsm', data_only=True)
    # sheet = wb['format']
    # EM0580106PData = pd.DataFrame(sheet.values)
    # EM0580106PData = EM0580106PData.replace(r'\s+', '', regex=True)

    EM0580106PData = pd.read_excel('EM0580106P NEW.xlsm', sheet_name='format', engine='calamine')
    EM0580106PData = EM0580106PData.replace(r'\s+', '', regex=True)

def ReadEM0580107P():
    global EM0580107PData

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    vt1Directory = (fr'\\192.168.2.19\quality control\2024\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\DHYE\~New Trend 2024')
    os.chdir(vt1Directory)

    # wb = load_workbook(filename='EM0580107P NEW.xlsm', data_only=True)
    # sheet = wb['format']
    # EM0580107PData = pd.DataFrame(sheet.values)
    # EM0580107PData = EM0580107PData.replace(r'\s+', '', regex=True)

    EM0580107PData = pd.read_excel('EM0580107P NEW.xlsm', sheet_name='format', engine='calamine')
    EM0580107PData = EM0580107PData.replace(r'\s+', '', regex=True)

def ReadFM05000102():
    global FM05000102Data

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    vt1Directory = (fr'\\192.168.2.19\quality control\2024\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\CRONICS\~DATA TREND 2024')
    os.chdir(vt1Directory)

    # wb = load_workbook(filename='FM05000102 NEW.xlsm', data_only=True)
    # sheet = wb['format']
    # FM05000102Data = pd.DataFrame(sheet.values)
    # FM05000102Data = FM05000102Data.replace(r'\s+', '', regex=True)

    FM05000102Data = pd.read_excel('FM05000102 NEW.xlsm', sheet_name='format', engine='calamine')
    FM05000102Data = FM05000102Data.replace(r'\s+', '', regex=True)

def GettingEM6P(lotNumber):
    global EM0580106PData
    
    global em6PTotalAverage3
    global em6PTotalAverage4
    global em6PTotalAverage5
    global em6PTotalAverage10

    global em6PTotalMinimum3
    global em6PTotalMinimum4
    global em6PTotalMinimum5

    global em6PTotalMaximum3
    global em6PTotalMaximum4
    global em6PTotalMaximum5

    global inspectionData

    average3 = 0
    average4 = 0
    average5 = 0
    em6PAverage10 = 0

    em6PTotalAverage3 = []
    em6PTotalAverage4 = []
    em6PTotalAverage5 = []
    em6PTotalAverage10 = []

    em6PTotalMinimum3 = []
    em6PTotalMinimum4 = []
    em6PTotalMinimum5 = []
    
    em6PTotalMaximum3 = []
    em6PTotalMaximum4 = []
    em6PTotalMaximum5 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in EM0580106PData.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = EM0580106PData.iloc[max(0, hiblowRow[0] - 3):min(len(EM0580106PData), hiblowRow[0] + 10), EM0580106PData.columns.get_loc(hiblowColumn[0]):EM0580106PData.columns.get_loc(hiblowColumn[0]) + 999]
    hiblowFiltered

    try:
        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print("Row indices:", lotNumberRow)
        print("Column names:", lotNumberColumn)

        inspectionData = EM0580106PData.iloc[max(0, lotNumberRow[0]):min(len(EM0580106PData), lotNumberRow[0] + 10), EM0580106PData.columns.get_loc(lotNumberColumn[0]):EM0580106PData.columns.get_loc(lotNumberColumn[0]) + 5]

        for a in range(0, len(lotNumberColumn)):
            # Get The Neighboring Data of Lot Number
            inspectionData = EM0580106PData.iloc[max(0, lotNumberRow[a]):min(len(EM0580106PData), lotNumberRow[a] + 10), EM0580106PData.columns.get_loc(lotNumberColumn[a]):EM0580106PData.columns.get_loc(lotNumberColumn[a]) + 5]

            average3 = inspectionData.iloc[5].mean()
            average4 = inspectionData.iloc[6].mean()
            average5 = inspectionData.iloc[7].mean()

            minimum3 = inspectionData.iloc[5].min()
            minimum4 = inspectionData.iloc[6].min()
            minimum5 = inspectionData.iloc[7].min()

            maximum3 = inspectionData.iloc[5].max()
            maximum4 = inspectionData.iloc[6].max()
            maximum5 = inspectionData.iloc[7].max()

            em6PTotalAverage3.append(average3)
            em6PTotalAverage4.append(average4)
            em6PTotalAverage5.append(average5)

            em6PTotalMinimum3.append(minimum3)
            em6PTotalMinimum4.append(minimum4)
            em6PTotalMinimum5.append(minimum5)

            em6PTotalMaximum3.append(maximum3)
            em6PTotalMaximum4.append(maximum4)
            em6PTotalMaximum5.append(maximum5)
        
        em6PTotalAverage3 = statistics.mean(em6PTotalAverage3)
        em6PTotalAverage4 = statistics.mean(em6PTotalAverage4)
        em6PTotalAverage5 = statistics.mean(em6PTotalAverage5)

        em6PTotalMinimum3 = min(em6PTotalMinimum3)
        em6PTotalMinimum4 = min(em6PTotalMinimum4)
        em6PTotalMinimum5 = min(em6PTotalMinimum5)

        em6PTotalMaximum3 = max(em6PTotalMaximum3)
        em6PTotalMaximum4 = max(em6PTotalMaximum4)
        em6PTotalMaximum5 = max(em6PTotalMaximum5)

        em6PTotalAverage3 = f"{em6PTotalAverage3:.2f}"
        em6PTotalAverage4 = f"{em6PTotalAverage4:.2f}"
        em6PTotalAverage5 = f"{em6PTotalAverage5:.2f}"

        em6PTotalMinimum3 = f"{em6PTotalMinimum3:.2f}"
        em6PTotalMinimum4 = f"{em6PTotalMinimum4:.2f}"
        em6PTotalMinimum5 = f"{em6PTotalMinimum5:.2f}"

        em6PTotalMaximum3 = f"{em6PTotalMaximum3:.2f}"
        em6PTotalMaximum4 = f"{em6PTotalMaximum4:.2f}"
        em6PTotalMaximum5 = f"{em6PTotalMaximum5:.2f}"

        print(f"Total Average: {em6PTotalAverage3}")
        print(f"Total Average: {em6PTotalAverage4}")
        print(f"Total Average: {em6PTotalAverage5}")

        print(f"Total Minimum: {em6PTotalMinimum3}")
        print(f"Total Minimum: {em6PTotalMinimum4}")
        print(f"Total Minimum: {em6PTotalMinimum5}")

        print(f"Total Maximum: {em6PTotalMaximum3}")
        print(f"Total Maximum: {em6PTotalMaximum4}")
        print(f"Total Maximum: {em6PTotalMaximum5}")
    
    except:
        em6PTotalAverage3 = 0
        em6PTotalAverage4 = 0
        em6PTotalAverage5 = 0

        em6PTotalMinimum3 = 0
        em6PTotalMinimum4 = 0
        em6PTotalMinimum5 = 0

        em6PTotalMaximum3 = 0
        em6PTotalMaximum4 = 0
        em6PTotalMaximum5 = 0

        print(f"Total Average: {em6PTotalAverage3}")
        print(f"Total Average: {em6PTotalAverage4}")
        print(f"Total Average: {em6PTotalAverage5}")

        print(f"Total Minimum: {em6PTotalMinimum3}")
        print(f"Total Minimum: {em6PTotalMinimum4}")
        print(f"Total Minimum: {em6PTotalMinimum5}")

        print(f"Total Maximum: {em6PTotalMaximum3}")
        print(f"Total Maximum: {em6PTotalMaximum4}")
        print(f"Total Maximum: {em6PTotalMaximum5}")

    #Getting The Row, Column Location Of SUPPLIER
    findSupplier = [(index, column) for index, row in EM0580106PData.iterrows() for column, value in row.items() if value == "SUPPLIER"]
    supplierRow = [index for index, _ in findSupplier]
    supplierColumn = [column for _, column in findSupplier]

    # Get the Neighboring Data Of Supplier
    supplierFiltered = EM0580106PData.iloc[max(0, supplierRow[0] - 3):min(len(EM0580106PData), supplierRow[0] + 10), EM0580106PData.columns.get_loc(supplierColumn[0]):EM0580106PData.columns.get_loc(supplierColumn[0]) + 999]
    print(supplierFiltered)

    # #Getting The Row, Column Location Of Lot Number
    # findSupplierLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    # supplierLotNumberRow = [index for index, _ in findSupplierLotNumber]
    # supplierLotNumberColumn = [column for _, column in findSupplierLotNumber]

    # for b in range(0, len(supplierLotNumberColumn)):
    #     supplierLotNumberFiltered = EM0580106PData.iloc[max(0, supplierLotNumberRow[b]):min(len(EM0580106PData), supplierLotNumberRow[b] + 13), EM0580106PData.columns.get_loc(supplierLotNumberColumn[b]):EM0580106PData.columns.get_loc(supplierLotNumberColumn[b]) + 5]

    #     em6PAverage10 = supplierLotNumberFiltered.iloc[12, 0]
    #     em6PTotalAverage10.append(em6PAverage10)

    # em6PTotalAverage10 = statistics.mean(em6PTotalAverage10)
    # em6PTotalAverage10 = f"{em6PTotalAverage10:.2f}"
    
    print(f"Total Average: {em6PTotalAverage10}")

def GettingEM7P(lotNumber):
    global EM0580107PData
    
    global em7PTotalAverage3
    global em7PTotalAverage4
    global em7PTotalAverage5

    global em7PTotalAverage10

    global inspectionData

    average3 = 0
    average4 = 0
    average5 = 0

    average10 = 0

    em7PTotalAverage3 = []
    em7PTotalAverage4 = []
    em7PTotalAverage5 = []

    em7PTotalAverage10 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in EM0580107PData.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = EM0580107PData.iloc[max(0, hiblowRow[0] - 3):min(len(EM0580107PData), hiblowRow[0] + 10), EM0580107PData.columns.get_loc(hiblowColumn[0]):EM0580107PData.columns.get_loc(hiblowColumn[0]) + 999]
    
    # try:
    #Getting The Row, Column Location Of Lot Number
    findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    lotNumberRow = [index for index, _ in findLotNumber]
    lotNumberColumn = [column for _, column in findLotNumber]


    for a in range(0, len(lotNumberColumn)):
        # Get The Neighboring Data of Lot Number
        hiblowLotNumberFiltered = EM0580107PData.iloc[max(0, lotNumberRow[a]):min(len(EM0580107PData), lotNumberRow[a] + 10), EM0580107PData.columns.get_loc(lotNumberColumn[a]):EM0580107PData.columns.get_loc(lotNumberColumn[a]) + 5]

        average3 = hiblowLotNumberFiltered.iloc[5].mean()
        average4 = hiblowLotNumberFiltered.iloc[6].mean()
        average5 = hiblowLotNumberFiltered.iloc[7].mean()

        em7PTotalAverage3.append(average3)
        em7PTotalAverage4.append(average4)
        em7PTotalAverage5.append(average5)
    
    em7PTotalAverage3 = statistics.mean(em7PTotalAverage3)
    em7PTotalAverage4 = statistics.mean(em7PTotalAverage4)
    em7PTotalAverage5 = statistics.mean(em7PTotalAverage5)

    em7PTotalAverage3 = f"{em7PTotalAverage3:.2f}"
    em7PTotalAverage4 = f"{em7PTotalAverage4:.2f}"
    em7PTotalAverage5 = f"{em7PTotalAverage5:.2f}"

    # #Getting The Row, Column Location Of SUPPLIER
    # findSupplier = [(index, column) for index, row in EM0580107PData.iterrows() for column, value in row.items() if value == "SUPPLIER"]
    # supplierRow = [index for index, _ in findSupplier]
    # supplierColumn = [column for _, column in findSupplier]

    # # Get the Neighboring Data Of Supplier
    # supplierFiltered = EM0580107PData.iloc[max(0, supplierRow[0] - 3):min(len(EM0580107PData), supplierRow[0] + 10), EM0580107PData.columns.get_loc(supplierColumn[0]):EM0580107PData.columns.get_loc(supplierColumn[0]) + 999]

    # #Getting The Row, Column Location Of Lot Number
    # findSupplierLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    # supplierLotNumberRow = [index for index, _ in findSupplierLotNumber]
    # supplierLotNumberColumn = [column for _, column in findSupplierLotNumber]

    # for b in range(0, len(supplierLotNumberColumn)):
    #     supplierLotNumberFiltered = EM0580107PData.iloc[max(0, supplierLotNumberRow[b]):min(len(EM0580107PData), supplierLotNumberRow[b] + 13), EM0580107PData.columns.get_loc(supplierLotNumberColumn[b]):EM0580107PData.columns.get_loc(supplierLotNumberColumn[b]) + 5]

    #     average10 = supplierLotNumberFiltered.iloc[12, 0]
    #     em7PTotalAverage10.append(average10)

    # em7PTotalAverage10 = statistics.mean(em7PTotalAverage10)
    # em7PTotalAverage10 = f"{em7PTotalAverage10:.2f}"

    # except:
    #     em7PTotalAverage3 = 0
    #     em7PTotalAverage4 = 0
    #     em7PTotalAverage5 = 0
    #     em7PTotalAverage10 = 0

    #     print(em7PTotalAverage3)
    #     print(em7PTotalAverage4)
    #     print(em7PTotalAverage5)
    #     print(em7PTotalAverage10)



def GettingFM(lotNumber):
    global FM05000102Data
    
    global totalAverage1
    global totalAverage2
    global totalAverage3
    global totalAverage4
    global totalAverage5
    global totalAverage6
    global totalAverage7

    global fmTotalMinimum1
    global fmTotalMinimum2
    global fmTotalMinimum3
    global fmTotalMinimum4
    global fmTotalMinimum5
    global fmTotalMinimum6
    global fmTotalMinimum7

    global inspectionData

    average1 = 0
    average2 = 0
    average3 = 0
    average4 = 0
    average5 = 0
    average6 = 0
    average7 = 0

    minimum1 = 0
    minimum2 = 0
    minimum3 = 0
    minimum4 = 0
    minimum5 = 0
    minimum6 = 0
    minimum7 = 0

    totalAverage1 = []
    totalAverage2 = []
    totalAverage3 = []
    totalAverage4 = []
    totalAverage5 = []
    totalAverage6 = []
    totalAverage7 = []

    fmTotalMinimum1 = []
    fmTotalMinimum2 = []
    fmTotalMinimum3 = []
    fmTotalMinimum4 = []
    fmTotalMinimum5 = []
    fmTotalMinimum6 = []
    fmTotalMinimum7 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in FM05000102Data.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = FM05000102Data.iloc[max(0, hiblowRow[0] - 3):min(len(FM05000102Data), hiblowRow[0] + 10), FM05000102Data.columns.get_loc(hiblowColumn[0]):FM05000102Data.columns.get_loc(hiblowColumn[0]) + 999]
    hiblowFiltered

    # try:
    #Getting The Row, Column Location Of Lot Number
    findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    lotNumberRow = [index for index, _ in findLotNumber]
    lotNumberColumn = [column for _, column in findLotNumber]

    print("Row indices:", lotNumberRow)
    print("Column names:", lotNumberColumn)


    for a in range(0, len(lotNumberColumn)):
        # Get The Neighboring Data of Lot Number
        inspectionData = FM05000102Data.iloc[max(0, lotNumberRow[a]):min(len(FM05000102Data), lotNumberRow[a] + 10), FM05000102Data.columns.get_loc(lotNumberColumn[a]):FM05000102Data.columns.get_loc(lotNumberColumn[a]) + 5]

        average1 = inspectionData.iloc[3].mean()
        average2 = inspectionData.iloc[4].mean()
        average3 = inspectionData.iloc[5].mean()
        average4 = inspectionData.iloc[6].mean()
        average5 = inspectionData.iloc[7].mean()
        average6 = inspectionData.iloc[8].mean()
        average7 = inspectionData.iloc[9].mean()

        minimum1 = inspectionData.iloc[3].min()
        minimum2 = inspectionData.iloc[4].min()
        minimum3 = inspectionData.iloc[5].min()
        minimum4 = inspectionData.iloc[6].min()
        minimum5 = inspectionData.iloc[7].min()
        minimum6 = inspectionData.iloc[8].min()
        minimum7 = inspectionData.iloc[9].min()

        totalAverage1.append(average1)
        totalAverage2.append(average2)
        totalAverage3.append(average3)
        totalAverage4.append(average4)
        totalAverage5.append(average5)
        totalAverage6.append(average6)
        totalAverage7.append(average7)

        fmTotalMinimum1.append(minimum1)
        fmTotalMinimum2.append(minimum2)
        fmTotalMinimum3.append(minimum3)
        fmTotalMinimum4.append(minimum4)
        fmTotalMinimum5.append(minimum5)
        fmTotalMinimum6.append(minimum6)
        fmTotalMinimum7.append(minimum7)

        inspectionData

    totalAverage1 = statistics.mean(totalAverage1)
    totalAverage2 = statistics.mean(totalAverage2)
    totalAverage3 = statistics.mean(totalAverage3)
    totalAverage4 = statistics.mean(totalAverage4)
    totalAverage5 = statistics.mean(totalAverage5)
    totalAverage6 = statistics.mean(totalAverage6)
    totalAverage7 = statistics.mean(totalAverage7)

    fmTotalMinimum1 = min(fmTotalMinimum1)
    fmTotalMinimum2 = min(fmTotalMinimum2)
    fmTotalMinimum3 = min(fmTotalMinimum3)
    fmTotalMinimum4 = min(fmTotalMinimum4)
    fmTotalMinimum5 = min(fmTotalMinimum5)
    fmTotalMinimum6 = min(fmTotalMinimum6)
    fmTotalMinimum7 = min(fmTotalMinimum7)

    totalAverage1 = f"{totalAverage1:.2f}"
    totalAverage2 = f"{totalAverage2:.2f}"
    totalAverage3 = f"{totalAverage3:.2f}"
    totalAverage4 = f"{totalAverage4:.2f}"
    totalAverage5 = f"{totalAverage5:.2f}"
    totalAverage6 = f"{totalAverage6:.2f}"
    totalAverage7 = f"{totalAverage7:.2f}"

    fmTotalMinimum1 = f"{fmTotalMinimum1:.2f}"
    fmTotalMinimum2 = f"{fmTotalMinimum2:.2f}"
    fmTotalMinimum3 = f"{fmTotalMinimum3:.2f}"
    fmTotalMinimum4 = f"{fmTotalMinimum4:.2f}"
    fmTotalMinimum5 = f"{fmTotalMinimum5:.2f}"
    fmTotalMinimum6 = f"{fmTotalMinimum6:.2f}"
    fmTotalMinimum7 = f"{fmTotalMinimum7:.2f}"

    print(totalAverage1)
    print(totalAverage2)
    print(totalAverage3)
    print(totalAverage4)
    print(totalAverage5)
    print(totalAverage6)
    print(totalAverage7)
    # except:
    #     totalAverage1 = 0
    #     totalAverage2 = 0
    #     totalAverage3 = 0
    #     totalAverage4 = 0
    #     totalAverage5 = 0
    #     totalAverage6 = 0
    #     totalAverage7 = 0
    #     print(totalAverage1)
    #     print(totalAverage2)
    #     print(totalAverage3)
    #     print(totalAverage4)
    #     print(totalAverage5)
    #     print(totalAverage6)
    #     print(totalAverage7)




# ReadFM()
# GettingFM("080224A-40")
# ReadEM6P()
# GettingEM6P("CAT-4H13DI")
# ReadEM0580107P()
# GettingEM7P("CAT-4A03DI")
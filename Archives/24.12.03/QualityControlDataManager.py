# %%
from Imports import *
import DateAndTimeManager

# %%
EM0580106PData = ""
EM0580107PData = ""
FM05000102Data = ""
CSB6400802Data = ""
DFB6600600Data = ""
RDB5200200Data = ""

inspectionData = ""

#EM2P
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

#EM3P
em7PTotalAverage3 = []
em7PTotalAverage4 = []
em7PTotalAverage5 = []
em7PTotalAverage10 = []

em7PTotalMinimum3 = []
em7PTotalMinimum4 = []
em7PTotalMinimum5 = []

em7PTotalMaximum3 = []
em7PTotalMaximum4 = []
em7PTotalMaximum5 = []

#FRAME
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

fmTotalMaximum1 = []
fmTotalMaximum2 = []
fmTotalMaximum3 = []
fmTotalMaximum4 = []
fmTotalMaximum5 = []
fmTotalMaximum6 = []
fmTotalMaximum7 = []

#CSB
csbTotalAverage1 = []

csbTotalMinimum1 = []

csbTotalMaximum1 = []

#DFB
dfbSnapData = ""
dfbLetterCode = ""
dfbLotNumber = ""
dfbMonth = ""
dfbCode = ""
dfbLotNumber2 = ""

dfbTotalAverage1 = []
dfbTotalAverage2 = []
dfbTotalAverage3 = []
dfbTotalAverage4 = []

dfbTotalMinimum1 = []
dfbTotalMinimum2 = []
dfbTotalMinimum3 = []
dfbTotalMinimum4 = []

dfbTotalMaximum1 = []
dfbTotalMaximum2 = []
dfbTotalMaximum3 = []
dfbTotalMaximum4 = []

#RDB
rdbCheckListData = ""
rdbLetterCode = ""
rdbMonth = ""
rdbLotNumber = ""
rdbLotNumber2 = ""
rdbLotNumber3 = ""
rdbCode = ""
rdbCode2 = ""
rdbProdDate = ""
rdbNoDataFound = ""

rdbTeslaTotalAverage1 = ""
rdbTeslaTotalAverage2 = ""
rdbTeslaTotalAverage3 = ""
rdbTeslaTotalAverage4 = ""

rdbTeslaTotalMinimum1 = ""
rdbTeslaTotalMinimum2 = ""
rdbTeslaTotalMinimum3 = ""
rdbTeslaTotalMinimum4 = ""

rdbTeslaTotalMaximum1 = ""
rdbTeslaTotalMaximum2 = ""
rdbTeslaTotalMaximum3 = ""
rdbTeslaTotalMaximum4 = ""

rdbTotalAverage1 = ""
rdbTotalAverage2 = ""
rdbTotalAverage3 = ""
rdbTotalAverage4 = ""
rdbTotalAverage5 = ""
rdbTotalAverage6 = ""
rdbTotalAverage8 = ""

rdbTotalMinimum1 = ""
rdbTotalMinimum2 = ""
rdbTotalMinimum3 = ""
rdbTotalMinimum4 = ""
rdbTotalMinimum5 = ""
rdbTotalMinimum6 = ""
rdbTotalMinimum8 = ""

rdbTotalMaximum1 = ""
rdbTotalMaximum2 = ""
rdbTotalMaximum3 = ""
rdbTotalMaximum4 = ""
rdbTotalMaximum5 = ""
rdbTotalMaximum6 = ""
rdbTotalMaximum8 = ""

# %%
def ReadEM0580106P():
    global EM0580106PData

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    try:
        vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\DHYE')

        #Finding A Folder That Contains New Trend
        for d in os.listdir(vt1Directory):
            if 'new trend' in d.lower():
                vt1Directory = os.path.join(vt1Directory, d)
                print(f"Updated vt1Directory: {vt1Directory}")
                break

        os.chdir(vt1Directory)

        # wb = load_workbook(filename='EM0580106P NEW.xlsm', data_only=True)
        # sheet = wb['format']
        # EM0580106PData = pd.DataFrame(sheet.values)
        # EM0580106PData = EM0580106PData.replace(r'\s+', '', regex=True)

        workbook = CalamineWorkbook.from_path("EM0580106P NEW.xlsm")
        EM0580106PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
        EM0580106PData = pd.DataFrame(EM0580106PData)
        EM0580106PData = EM0580106PData.replace(r'\s+', '', regex=True)
    except:
        try:
            #Finding A Folder That Contains Data Trend
            for d in os.listdir(vt1Directory):
                if 'data trend' in d.lower():
                    vt1Directory = os.path.join(vt1Directory, d)
                    print(f"Updated vt1Directory: {vt1Directory}")
                    break

            os.chdir(vt1Directory)

            # wb = load_workbook(filename='EM0580106P NEW.xlsm', data_only=True)
            # sheet = wb['format']
            # EM0580106PData = pd.DataFrame(sheet.values)
            # EM0580106PData = EM0580106PData.replace(r'\s+', '', regex=True)

            workbook = CalamineWorkbook.from_path("EM0580106P NEW.xlsm")
            EM0580106PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
            EM0580106PData = pd.DataFrame(EM0580106PData)
            EM0580106PData = EM0580106PData.replace(r'\s+', '', regex=True)
        except:
            print("Error: Unable to read EM0580106P data")

    # EM0580106PData = pd.read_excel('EM0580106P NEW.xlsm', sheet_name='format', engine='calamine')
    # EM0580106PData = EM0580106PData.replace(r'\s+', '', regex=True)

# %%
# DateAndTimeManager.GetDateToday()

# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

# vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\DHYE')
# for d in os.listdir(vt1Directory):
#     if 'new trend' in d.lower():
#         vt1Directory = os.path.join(vt1Directory, d)
#         print(f"Updated vt1Directory: {vt1Directory}")
#         break

# os.chdir(vt1Directory)

# workbook = CalamineWorkbook.from_path("EM0580106P NEW.xlsm")
# # sheet_names = workbook.sheet_names
# EM0580106PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)

# EM0580106PData = pd.DataFrame(EM0580106PData)

# EM0580106PData

# # ReadEM0580106P()

# %%
# DateAndTimeManager.GetDateToday()
# ReadEM0580106P()
# EM0580106PData

# %%
def ReadEM0580107P():
    global EM0580107PData

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    try:
        vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\DHYE')
        
        #Finding A Folder That Contains New Trend
        for d in os.listdir(vt1Directory):
            if 'new trend' in d.lower():
                vt1Directory = os.path.join(vt1Directory, d)
                print(f"Updated vt1Directory: {vt1Directory}")
                break
        
        os.chdir(vt1Directory)

        # wb = load_workbook(filename='EM0580107P NEW.xlsm', data_only=True)
        # sheet = wb['format']
        # EM0580107PData = pd.DataFrame(sheet.values)
        # EM0580107PData = EM0580107PData.replace(r'\s+', '', regex=True)

        workbook = CalamineWorkbook.from_path("EM0580107P NEW.xlsm")
        EM0580107PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=False)
        EM0580107PData = pd.DataFrame(EM0580107PData)
        # EM0580107PData = EM0580107PData.replace(r'\s+', '', regex=True)

        #Making Blank Values To None
        EM0580107PData.replace(r'^\s*$', np.nan, regex=True, inplace=True)
        
        EM0580107PData = EM0580107PData.where(pd.notnull(EM0580107PData), None)

    except:
        try:
            vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\DHYE')
            
            #Finding A Folder That Contains Data Trend
            for d in os.listdir(vt1Directory):
                if 'data trend' in d.lower():
                    vt1Directory = os.path.join(vt1Directory, d)
                    print(f"Updated vt1Directory: {vt1Directory}")
                    break
            
            os.chdir(vt1Directory)

            # wb = load_workbook(filename='EM0580107P NEW.xlsm', data_only=True)
            # sheet = wb['format']
            # EM0580107PData = pd.DataFrame(sheet.values)
            # EM0580107PData = EM0580107PData.replace(r'\s+', '', regex=True)

            workbook = CalamineWorkbook.from_path("EM0580107P NEW.xlsm")
            EM0580107PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=False)
            EM0580107PData = pd.DataFrame(EM0580107PData)
            # EM0580107PData = EM0580107PData.replace(r'\s+', '', regex=True)
            EM0580107PData.replace(r'^\s*$', np.nan, regex=True, inplace=True)
            EM0580107PData = EM0580107PData.where(pd.notnull(EM0580107PData), None)
        except:
            print("Error: Unable to read EM0580107P data.")

    # EM0580107PData = pd.read_excel('EM0580107P NEW.xlsm', sheet_name='format', engine='calamine')
    # EM0580107PData = EM0580107PData.replace(r'\s+', '', regex=True)

# %%
# DateAndTimeManager.GetDateToday()
# ReadEM0580107P()

# EM0580107PData

# %%
def ReadFM05000102():
    global FM05000102Data

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    try:
        vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\CRONICS')

        #Finding A Folder That Contains New Trend
        for d in os.listdir(vt1Directory):
            if 'new trend' in d.lower():
                vt1Directory = os.path.join(vt1Directory, d)
                print(f"Updated vt1Directory: {vt1Directory}")
                break
        
        os.chdir(vt1Directory)

        # wb = load_workbook(filename='FM05000102 NEW.xlsm', data_only=True)
        # sheet = wb['format']
        # FM05000102Data = pd.DataFrame(sheet.values)
        # FM05000102Data = FM05000102Data.replace(r'\s+', '', regex=True)

        workbook = CalamineWorkbook.from_path("FM05000102 NEW.xlsm")
        FM05000102Data = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
        FM05000102Data = pd.DataFrame(FM05000102Data)
        FM05000102Data = FM05000102Data.replace(r'\s+', '', regex=True)

    except:
        try:
            vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\CRONICS')

            #Finding A Folder That Contains Data Trend
            for d in os.listdir(vt1Directory):
                if 'data trend' in d.lower():
                    vt1Directory = os.path.join(vt1Directory, d)
                    print(f"Updated vt1Directory: {vt1Directory}")
                    break
            
            os.chdir(vt1Directory)

            # wb = load_workbook(filename='FM05000102 NEW.xlsm', data_only=True)
            # sheet = wb['format']
            # FM05000102Data = pd.DataFrame(sheet.values)
            # FM05000102Data = FM05000102Data.replace(r'\s+', '', regex=True)

            workbook = CalamineWorkbook.from_path("FM05000102 NEW.xlsm")
            FM05000102Data = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
            FM05000102Data = pd.DataFrame(FM05000102Data)
            FM05000102Data = FM05000102Data.replace(r'\s+', '', regex=True)
        except:
            print("Error: Unable to read FM05000102 data")

    # FM05000102Data = pd.read_excel('FM05000102 NEW.xlsm', sheet_name='format', engine='calamine')
    # FM05000102Data = FM05000102Data.replace(r'\s+', '', regex=True)

# %%
def ReadCSB6400802():
    global CSB6400802Data

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    
    try:
        vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\CRONICS')
        
        #Finding A Folder That Contains New Trend
        for d in os.listdir(vt1Directory):
            if 'new trend' in d.lower():
                vt1Directory = os.path.join(vt1Directory, d)
                print(f"Updated vt1Directory: {vt1Directory}")
                break
        
        os.chdir(vt1Directory)

        workbook = CalamineWorkbook.from_path("CSB6400802 NEW.xlsm")
        CSB6400802Data = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
        CSB6400802Data = pd.DataFrame(CSB6400802Data)
        CSB6400802Data = CSB6400802Data.replace(r'\s+', '', regex=True)

    except:
        try:
            vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\CRONICS')
            
            #Finding A Folder That Contains New Trend
            for d in os.listdir(vt1Directory):
                if 'data trend' in d.lower():
                    vt1Directory = os.path.join(vt1Directory, d)
                    print(f"Updated vt1Directory: {vt1Directory}")
                    break
            
            os.chdir(vt1Directory)

            workbook = CalamineWorkbook.from_path("CSB6400802 NEW.xlsm")
            CSB6400802Data = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
            CSB6400802Data = pd.DataFrame(CSB6400802Data)
            CSB6400802Data = CSB6400802Data.replace(r'\s+', '', regex=True)
        except:
            print('CSB6400802Data Not Found')

# %%
# DateAndTimeManager.GetDateToday()
# ReadCSB6400802()
# CSB6400802Data

# %%
def ReadDFBSnap(lotNumber):
    global dfbSnapData
    global dfbLetterCode
    global dfbLotNumber
    global dfbMonth

    dfbLetterCode = lotNumber[-1]


    #Removing The Last Two Values Of Lot Number
    lotNumber = lotNumber[:-2]
    #Changing The Format Of Lot Number
    lotNumber = datetime2.strptime(lotNumber, "%Y%m%d")
    dfbMonth = lotNumber.strftime("%B")
    dfbLotNumber = lotNumber.strftime("%Y-%m-%d")
    
    print(dfbLotNumber)
    print(dfbMonth)
    print(dfbLetterCode)

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    try:
        vt1Directory = (fr'\\192.168.2.19\production\{DateAndTimeManager.yearNow}\3. On-line Checksheet\OUTJOB\OUTJOB MATERIAL MONITORING CHECKSHEET')

        os.chdir(vt1Directory)

        wb = load_workbook(filename='SNAP.xlsx', data_only=True)

        for s in wb.sheetnames:
            if dfbMonth.lower() in s.lower():
                sheet = wb[s]
                dfbSnapData = pd.DataFrame(sheet.values)
                dfbSnapData = dfbSnapData.iloc[6:]
                dfbSnapData = dfbSnapData.replace(r'\s+', '', regex=True)

                break

    except:
        print('DFB6600600 Not Found')

# %%
def ReadDFB6600600():
    global dfbSnapData
    global dfbLetterCode
    global dfbLotNumber

    global dfbCode
    global dfbLotNumber2

    dfbCode = dfbSnapData.iloc[1, 3]
    dfbCode = dfbCode[8:]
    dfbCode = dfbCode[:-28]

    #Converting The First Column/Date To String
    dfbSnapData.iloc[:, 0] = dfbSnapData.iloc[:, 0].astype(str)

    tempDfbSnapData = dfbSnapData[(dfbSnapData[0].isin([f"{dfbLotNumber} 00:00:00"])) & (dfbSnapData[2].isin([dfbLetterCode]))]

    dfbLotNumber2 = tempDfbSnapData.iloc[:,3].values[0]

# %%
def ReadRDB5200200CheckSheet(lotNumber):
    global RDB5200200Data

    global rdbLetterCode
    global rdbMonth
    global rdbLotNumber
    global rdbLotNumber2
    global rdbLotNumber3
    global rdbCode
    global rdbCode2
    global rdbProdDate

    global rdbTeslaTotalAverage1
    global rdbTeslaTotalAverage2
    global rdbTeslaTotalAverage3
    global rdbTeslaTotalAverage4

    global rdbTeslaTotalMinimum1
    global rdbTeslaTotalMinimum2
    global rdbTeslaTotalMinimum3
    global rdbTeslaTotalMinimum4

    global rdbTeslaTotalMaximum1
    global rdbTeslaTotalMaximum2
    global rdbTeslaTotalMaximum3
    global rdbTeslaTotalMaximum4

    global rdbNoDataFound

    try:
        rdbLetterCode = lotNumber[-1]

        #Removing The Last Two Values Of Lot Number
        lotNumber = lotNumber[:-2]
        #Changing The Format Of Lot Number
        lotNumber = datetime2.strptime(lotNumber, "%Y%m%d")
        rdbMonth = lotNumber.strftime("%B")
        rdbLotNumber = lotNumber.strftime("%d/%m/%Y")

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        vt1Directory = (fr'\\192.168.2.19\production\{DateAndTimeManager.yearNow}\3. On-line Checksheet\OUTJOB\ROD CHECKSHEET')
        
        os.chdir(vt1Directory)

        #Finding All xlsm Files In The Current Directory
        files = glob.glob('*.xlsx')

        recentTime = 0

        #Checking Each Files In Files;
        for f in files:
            if 'RDB5200200' in f:
                #Checking If It Is Recent File
                fileTime = os.path.getmtime(f)
                if fileTime > recentTime:
                    recentTime = fileTime
                    fileName = f

        targetValue = rdbMonth

        workbook = CalamineWorkbook.from_path(fileName)
        for s in workbook.sheet_names:
            if targetValue.lower() in s.lower():
                RDB5200200Data = workbook.get_sheet_by_name(s).to_python(skip_empty_area=True)
                RDB5200200Data = pd.DataFrame(RDB5200200Data)
                RDB5200200Data = RDB5200200Data.replace(r'\s+', '', regex=True)
                break

        #Getting The RDB Code
        rdbCode = RDB5200200Data.iloc[4, 9]
        rdbCode = rdbCode[:-3]
        rdbCode = rdbCode[27:]
        #__________________________________
        #Getting The RDB Code 2
        rdbCode2 = RDB5200200Data.iloc[4, 9]
        rdbCode2 = rdbCode2[27:]
        #__________________________________

        #Getting The Lot Number (Supplier)
        rdbLotNumber2 = RDB5200200Data.iloc[max(0 + 7, 0):min(len(RDB5200200Data), 0 + 999), RDB5200200Data.columns.get_loc(0):RDB5200200Data.columns.get_loc(0) + 11]
        rdbLotNumber2 = rdbLotNumber2[(rdbLotNumber2[0].isin([rdbLotNumber])) & (rdbLotNumber2[8].isin([rdbLetterCode]))]
        rdbProdDate = rdbLotNumber2[10].values[0]
        rdbProdDate = rdbProdDate[:-1]
        rdbLotNumber2 = rdbLotNumber2[9].values[0]

        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in RDB5200200Data.iterrows() for column, value in row.items() if str(value) == str(rdbLotNumber)]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print("Row indices:", lotNumberRow)
        print("Column names:", lotNumberColumn)

        #Getting The Tesla Table
        inspectionData = RDB5200200Data.iloc[max(0, lotNumberRow[0]):min(len(RDB5200200Data), lotNumberRow[0] + 7), RDB5200200Data.columns.get_loc(lotNumberColumn[0] + 21):RDB5200200Data.columns.get_loc(lotNumberColumn[0]) + 26]

        rdbTeslaTotalAverage1 = inspectionData.iloc[0].mean()
        rdbTeslaTotalAverage2 = inspectionData.iloc[2].mean()
        rdbTeslaTotalAverage3 = inspectionData.iloc[4].mean()
        rdbTeslaTotalAverage4 = inspectionData.iloc[6].mean()

        rdbTeslaTotalMinimum1 = inspectionData.iloc[0].min()
        rdbTeslaTotalMinimum2 = inspectionData.iloc[2].min()
        rdbTeslaTotalMinimum3 = inspectionData.iloc[4].min()
        rdbTeslaTotalMinimum4 = inspectionData.iloc[6].min()

        rdbTeslaTotalMaximum1 = inspectionData.iloc[0].max()
        rdbTeslaTotalMaximum2 = inspectionData.iloc[2].max()
        rdbTeslaTotalMaximum3 = inspectionData.iloc[4].max()
        rdbTeslaTotalMaximum4 = inspectionData.iloc[6].max()
        
        print(f"Tesla Average 1:{rdbTeslaTotalAverage1}")
        print(f"Tesla Average 2:{rdbTeslaTotalAverage2}")
        print(f"Tesla Average 3:{rdbTeslaTotalAverage3}")
        print(f"Tesla Average 4:{rdbTeslaTotalAverage4}")

        print(f"Tesla Minimum 1:{rdbTeslaTotalMinimum1}")
        print(f"Tesla Minimum 2:{rdbTeslaTotalMinimum2}")
        print(f"Tesla Minimum 3:{rdbTeslaTotalMinimum3}")
        print(f"Tesla Minimum 4:{rdbTeslaTotalMinimum4}")

        print(f"Tesla Maximum 1:{rdbTeslaTotalMaximum1}")
        print(f"Tesla Maximum 2:{rdbTeslaTotalMaximum2}")
        print(f"Tesla Maximum 3:{rdbTeslaTotalMaximum3}")
        print(f"Tesla Maximum 4:{rdbTeslaTotalMaximum4}")

        #Reading HPIQCDATA
        hpiQcDataDirectory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\B. Monitoring Files')
        os.chdir(hpiQcDataDirectory)

        xlsxFiles = glob.glob('*.xlsx')
        xlsFiles = glob.glob('*.xls')

        files = xlsxFiles + xlsFiles

        #Checking Each Files In Files;
        for f in files:
            if 'HPI-QA'.lower() in f.lower() or "HPI-QC".lower() in f.lower():
                
                workbook = CalamineWorkbook.from_path(f)

                #Reading Possible Sheets
                try:
                    hpiQAQCData = workbook.get_sheet_by_name("HPI-QC01-01").to_python(skip_empty_area=True)
                    hpiQAQCData = pd.DataFrame(hpiQAQCData[1:], columns=hpiQAQCData[2])
                except:
                    hpiQAQCData = workbook.get_sheet_by_name("SUMMARY").to_python(skip_empty_area=True)
                    hpiQAQCData = pd.DataFrame(hpiQAQCData[1:], columns=hpiQAQCData[0])

                hpiQAQCData['DATE RECEIVED'] = hpiQAQCData['DATE RECEIVED'].astype(str).str.replace("-", "")
                hpiQAQCData = hpiQAQCData[(hpiQAQCData["DATE RECEIVED"].isin([str(rdbProdDate)])) & (hpiQAQCData["ITEM CODE"].isin([str(rdbCode2)]))]
                hpiQAQCData = hpiQAQCData[hpiQAQCData['LOT NUMBER'].str.contains(rdbLotNumber2[:-3], na=False)]


                if hpiQAQCData.empty:
                    print(f"No data found for {f}")
                else:
                    break

        rdbLotNumber3 = hpiQAQCData["LOT NUMBER"].values[0]
        rdbNoDataFound = False
    except:
        rdbNoDataFound = True
        print("No RDB Data Found")

# %%
def ReadRDB5200200():
    global RDB5200200Data
    global rdbCode

    global rdbNoDataFound

    if not rdbNoDataFound:

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        try:
            vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\SBROS')

            #Finding A Folder That Contains New Trend
            for d in os.listdir(vt1Directory):
                if 'new trend' in d.lower():
                    vt1Directory = os.path.join(vt1Directory, d)
                    print(f"Updated vt1Directory: {vt1Directory}")
                    break

            os.chdir(vt1Directory)

            workbook = CalamineWorkbook.from_path("RD05200200 NEW.xlsm")
            RDB5200200Data = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
            RDB5200200Data = pd.DataFrame(RDB5200200Data)
            RDB5200200Data = RDB5200200Data.replace(r'\s+', '', regex=True)
        except:
            try:
                #Finding A Folder That Contains Data Trend
                for d in os.listdir(vt1Directory):
                    if 'data trend' in d.lower():
                        vt1Directory = os.path.join(vt1Directory, d)
                        print(f"Updated vt1Directory: {vt1Directory}")
                        break

                os.chdir(vt1Directory)

                workbook = CalamineWorkbook.from_path("RD05200200 NEW.xlsm")
                RDB5200200Data = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                RDB5200200Data = pd.DataFrame(RDB5200200Data)
                RDB5200200Data = RDB5200200Data.replace(r'\s+', '', regex=True)
            except:
                print("Error: Unable to read RDB5200200 data")

    

# %%
def ReadAll():
    DateAndTimeManager.GetDateToday()
    ReadEM0580106P()
    ReadEM0580107P()
    ReadFM05000102()
    ReadCSB6400802()

# %%
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
    supplierFiltered

    #Getting The Row, Column Location Of Lot Number
    findSupplierLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    supplierLotNumberRow = [index for index, _ in findSupplierLotNumber]
    supplierLotNumberColumn = [column for _, column in findSupplierLotNumber]

    for b in range(0, len(supplierLotNumberColumn)):
        supplierLotNumberFiltered = EM0580106PData.iloc[max(0, supplierLotNumberRow[b]):min(len(EM0580106PData), supplierLotNumberRow[b] + 13), EM0580106PData.columns.get_loc(supplierLotNumberColumn[b]):EM0580106PData.columns.get_loc(supplierLotNumberColumn[b]) + 5]

        em6PAverage10 = supplierLotNumberFiltered.iloc[12, 0]
        em6PTotalAverage10.append(em6PAverage10)

    em6PTotalAverage10 = statistics.mean(em6PTotalAverage10)
    em6PTotalAverage10 = f"{em6PTotalAverage10:.2f}"
    
    print(f"Total Average: {em6PTotalAverage10}")

# %%
def GettingEM7P(lotNumber):
    global EM0580107PData
    
    global em7PTotalAverage3
    global em7PTotalAverage4
    global em7PTotalAverage5
    global em7PTotalAverage10

    global em7PTotalMinimum3
    global em7PTotalMinimum4
    global em7PTotalMinimum5

    global em7PTotalMaximum3
    global em7PTotalMaximum4
    global em7PTotalMaximum5

    global inspectionData

    average3 = 0
    average4 = 0
    average5 = 0

    average10 = 0

    em7PTotalAverage3 = []
    em7PTotalAverage4 = []
    em7PTotalAverage5 = []
    em7PTotalAverage10 = []

    em7PTotalMinimum3 = []
    em7PTotalMinimum4 = []
    em7PTotalMinimum5 = []

    em7PTotalMaximum3 = []
    em7PTotalMaximum4 = []
    em7PTotalMaximum5 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in EM0580107PData.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = EM0580107PData.iloc[max(0, hiblowRow[0] - 3):min(len(EM0580107PData), hiblowRow[0] + 10), EM0580107PData.columns.get_loc(hiblowColumn[0]):EM0580107PData.columns.get_loc(hiblowColumn[0]) + 999]
    
    try:
        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print(lotNumberRow)
        print(lotNumberColumn)

        for a in range(0, len(lotNumberColumn)):
            # Get The Neighboring Data of Lot Number
            hiblowLotNumberFiltered = EM0580107PData.iloc[max(0, lotNumberRow[a]):min(len(EM0580107PData), lotNumberRow[a] + 10), EM0580107PData.columns.get_loc(lotNumberColumn[a]):EM0580107PData.columns.get_loc(lotNumberColumn[a]) + 5]

            print("Correct")

            average3 = hiblowLotNumberFiltered.iloc[5].mean()
            average4 = hiblowLotNumberFiltered.iloc[6].mean()
            average5 = hiblowLotNumberFiltered.iloc[7].mean()

            minimum3 = hiblowLotNumberFiltered.iloc[5].mean()
            minimum4 = hiblowLotNumberFiltered.iloc[6].mean()
            minimum5 = hiblowLotNumberFiltered.iloc[7].mean()
            
            maximum3 = hiblowLotNumberFiltered.iloc[5].mean()
            maximum4 = hiblowLotNumberFiltered.iloc[6].mean()
            maximum5 = hiblowLotNumberFiltered.iloc[7].mean()

            em7PTotalAverage3.append(average3)
            em7PTotalAverage4.append(average4)
            em7PTotalAverage5.append(average5)

            em7PTotalMinimum3.append(minimum3)
            em7PTotalMinimum4.append(minimum4)
            em7PTotalMinimum5.append(minimum5)

            em7PTotalMaximum3.append(maximum3)
            em7PTotalMaximum4.append(maximum4)
            em7PTotalMaximum5.append(maximum5)
        
        em7PTotalAverage3 = statistics.mean(em7PTotalAverage3)
        em7PTotalAverage4 = statistics.mean(em7PTotalAverage4)
        em7PTotalAverage5 = statistics.mean(em7PTotalAverage5)

        em7PTotalMinimum3 = min(em7PTotalMinimum3)
        em7PTotalMinimum4 = min(em7PTotalMinimum4)
        em7PTotalMinimum5 = min(em7PTotalMinimum5)

        em7PTotalMaximum3 = max(em7PTotalMaximum3)
        em7PTotalMaximum4 = max(em7PTotalMaximum4)
        em7PTotalMaximum5 = max(em7PTotalMaximum5)

        em7PTotalAverage3 = f"{em7PTotalAverage3:.2f}"
        em7PTotalAverage4 = f"{em7PTotalAverage4:.2f}"
        em7PTotalAverage5 = f"{em7PTotalAverage5:.2f}"
        
        em7PTotalMinimum3 = f"{em7PTotalMinimum3:.2f}"
        em7PTotalMinimum4 = f"{em7PTotalMinimum4:.2f}"
        em7PTotalMinimum5 = f"{em7PTotalMinimum5:.2f}"

        em7PTotalMaximum3 = f"{em7PTotalMaximum3:.2f}"
        em7PTotalMaximum4 = f"{em7PTotalMaximum4:.2f}"
        em7PTotalMaximum5 = f"{em7PTotalMaximum5:.2f}"

        print(f"Total Average: {em7PTotalAverage3}")
        print(f"Total Average: {em7PTotalAverage4}")
        print(f"Total Average: {em7PTotalAverage5}")

        print(f"Total Minimum: {em7PTotalMinimum3}")
        print(f"Total Minimum: {em7PTotalMinimum4}")
        print(f"Total Minimum: {em7PTotalMinimum5}")

        print(f"Total Maximum: {em7PTotalMaximum3}")
        print(f"Total Maximum: {em7PTotalMaximum4}")
        print(f"Total Maximum: {em7PTotalMaximum5}")

    except:
        em7PTotalAverage3 = 0
        em7PTotalAverage4 = 0
        em7PTotalAverage5 = 0

        em7PTotalMinimum3 = 0
        em7PTotalMinimum4 = 0
        em7PTotalMinimum5 = 0

        em7PTotalMaximum3 = 0
        em7PTotalMaximum4 = 0
        em7PTotalMaximum5 = 0
        # em7PTotalAverage10 = 0

        print(f"Total Average: {em7PTotalAverage3}")
        print(f"Total Average: {em7PTotalAverage4}")
        print(f"Total Average: {em7PTotalAverage5}")

        print(f"Total Minimum: {em7PTotalMinimum3}")
        print(f"Total Minimum: {em7PTotalMinimum4}")
        print(f"Total Minimum: {em7PTotalMinimum5}")

        print(f"Total Maximum: {em7PTotalMaximum3}")
        print(f"Total Maximum: {em7PTotalMaximum4}")
        print(f"Total Maximum: {em7PTotalMaximum5}")

    #Getting The Row, Column Location Of SUPPLIER
    findSupplier = [(index, column) for index, row in EM0580107PData.iterrows() for column, value in row.items() if value == "SUPPLIER"]
    supplierRow = [index for index, _ in findSupplier]
    supplierColumn = [column for _, column in findSupplier]

    # Get the Neighboring Data Of Supplier
    supplierFiltered = EM0580107PData.iloc[max(0, supplierRow[0] - 3):min(len(EM0580107PData), supplierRow[0] + 10), EM0580107PData.columns.get_loc(supplierColumn[0]):EM0580107PData.columns.get_loc(supplierColumn[0]) + 999]

    #Getting The Row, Column Location Of Lot Number
    findSupplierLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
    supplierLotNumberRow = [index for index, _ in findSupplierLotNumber]
    supplierLotNumberColumn = [column for _, column in findSupplierLotNumber]

    for b in range(0, len(supplierLotNumberColumn)):
        supplierLotNumberFiltered = EM0580107PData.iloc[max(0, supplierLotNumberRow[b]):min(len(EM0580107PData), supplierLotNumberRow[b] + 13), EM0580107PData.columns.get_loc(supplierLotNumberColumn[b]):EM0580107PData.columns.get_loc(supplierLotNumberColumn[b]) + 5]

        average10 = supplierLotNumberFiltered.iloc[12, 0]
        em7PTotalAverage10.append(average10)

    #Removing None Values In Array
    while None in em7PTotalAverage10:
        em7PTotalAverage10.remove(None)
    #_____________________________

    em7PTotalAverage10 = statistics.mean(em7PTotalAverage10)
    em7PTotalAverage10 = f"{em7PTotalAverage10:.2f}"

# %%
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

    global fmTotalMaximum1
    global fmTotalMaximum2
    global fmTotalMaximum3
    global fmTotalMaximum4
    global fmTotalMaximum5
    global fmTotalMaximum6
    global fmTotalMaximum7

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

    fmTotalMaximum1 = []
    fmTotalMaximum2 = []
    fmTotalMaximum3 = []
    fmTotalMaximum4 = []
    fmTotalMaximum5 = []
    fmTotalMaximum6 = []
    fmTotalMaximum7 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in FM05000102Data.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = FM05000102Data.iloc[max(0, hiblowRow[0] - 3):min(len(FM05000102Data), hiblowRow[0] + 10), FM05000102Data.columns.get_loc(hiblowColumn[0]):FM05000102Data.columns.get_loc(hiblowColumn[0]) + 999]
    hiblowFiltered

    try:
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

            maximum1 = inspectionData.iloc[3].max()
            maximum2 = inspectionData.iloc[4].max()
            maximum3 = inspectionData.iloc[5].max()
            maximum4 = inspectionData.iloc[6].max()
            maximum5 = inspectionData.iloc[7].max()
            maximum6 = inspectionData.iloc[8].max()
            maximum7 = inspectionData.iloc[9].max()

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

            fmTotalMaximum1.append(maximum1)
            fmTotalMaximum2.append(maximum2)
            fmTotalMaximum3.append(maximum3)
            fmTotalMaximum4.append(maximum4)
            fmTotalMaximum5.append(maximum5)
            fmTotalMaximum6.append(maximum6)
            fmTotalMaximum7.append(maximum7)

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

        fmTotalMaximum1 = max(fmTotalMaximum1)
        fmTotalMaximum2 = max(fmTotalMaximum2)
        fmTotalMaximum3 = max(fmTotalMaximum3)
        fmTotalMaximum4 = max(fmTotalMaximum4)
        fmTotalMaximum5 = max(fmTotalMaximum5)
        fmTotalMaximum6 = max(fmTotalMaximum6)
        fmTotalMaximum7 = max(fmTotalMaximum7)

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

        fmTotalMaximum1 = f"{fmTotalMaximum1:.2f}"
        fmTotalMaximum2 = f"{fmTotalMaximum2:.2f}"
        fmTotalMaximum3 = f"{fmTotalMaximum3:.2f}"
        fmTotalMaximum4 = f"{fmTotalMaximum4:.2f}"
        fmTotalMaximum5 = f"{fmTotalMaximum5:.2f}"
        fmTotalMaximum6 = f"{fmTotalMaximum6:.2f}"
        fmTotalMaximum7 = f"{fmTotalMaximum7:.2f}"

        print(f"Average:{totalAverage1}")
        print(f"Average:{totalAverage2}")
        print(f"Average:{totalAverage3}")
        print(f"Average:{totalAverage4}")
        print(f"Average:{totalAverage5}")
        print(f"Average:{totalAverage6}")
        print(f"Average:{totalAverage7}")

        print(f"Minimum:{fmTotalMinimum1}")
        print(f"Minimum:{fmTotalMinimum2}")
        print(f"Minimum:{fmTotalMinimum3}")
        print(f"Minimum:{fmTotalMinimum4}")
        print(f"Minimum:{fmTotalMinimum5}")
        print(f"Minimum:{fmTotalMinimum6}")
        print(f"Minimum:{fmTotalMinimum7}")

        print(f"Maximum:{fmTotalMaximum1}")
        print(f"Maximum:{fmTotalMaximum2}")
        print(f"Maximum:{fmTotalMaximum3}")
        print(f"Maximum:{fmTotalMaximum4}")
        print(f"Maximum:{fmTotalMaximum5}")
        print(f"Maximum:{fmTotalMaximum6}")
        print(f"Maximum:{fmTotalMaximum7}")
    except:
        totalAverage1 = 0
        totalAverage2 = 0
        totalAverage3 = 0
        totalAverage4 = 0
        totalAverage5 = 0
        totalAverage6 = 0
        totalAverage7 = 0

        minimum1 = 0
        minimum2 = 0
        minimum3 = 0
        minimum4 = 0
        minimum5 = 0
        minimum6 = 0
        minimum7 = 0

        maximum1 = 0
        maximum2 = 0
        maximum3 = 0
        maximum4 = 0
        maximum5 = 0
        maximum6 = 0
        maximum7 = 0

        print(f"Average:{totalAverage1}")
        print(f"Average:{totalAverage2}")
        print(f"Average:{totalAverage3}")
        print(f"Average:{totalAverage4}")
        print(f"Average:{totalAverage5}")
        print(f"Average:{totalAverage6}")
        print(f"Average:{totalAverage7}")

        print(f"Minimum:{fmTotalMinimum1}")
        print(f"Minimum:{fmTotalMinimum2}")
        print(f"Minimum:{fmTotalMinimum3}")
        print(f"Minimum:{fmTotalMinimum4}")
        print(f"Minimum:{fmTotalMinimum5}")
        print(f"Minimum:{fmTotalMinimum6}")
        print(f"Minimum:{fmTotalMinimum7}")

        print(f"Maximum:{fmTotalMaximum1}")
        print(f"Maximum:{fmTotalMaximum2}")
        print(f"Maximum:{fmTotalMaximum3}")
        print(f"Maximum:{fmTotalMaximum4}")
        print(f"Maximum:{fmTotalMaximum5}")
        print(f"Maximum:{fmTotalMaximum6}")
        print(f"Maximum:{fmTotalMaximum7}")

# %%
def GettingCSB6400802(lotNumber):
    global CSB6400802Data

    global csbTotalAverage1

    global csbTotalMinimum1

    global csbTotalMaximum1

    csbTotalAverage1 = []

    csbTotalMinimum1 = []

    csbTotalMaximum1 = []

    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in CSB6400802Data.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = CSB6400802Data.iloc[max(0, hiblowRow[0] - 3):min(len(CSB6400802Data), hiblowRow[0] + 10), CSB6400802Data.columns.get_loc(hiblowColumn[0]):CSB6400802Data.columns.get_loc(hiblowColumn[0]) + 999]

    try:
        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == lotNumber]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print("Row indices:", lotNumberRow)
        print("Column names:", lotNumberColumn)

        for a in range(0, len(lotNumberColumn)):
            # Get The Neighboring Data of Lot Number
            inspectionData = CSB6400802Data.iloc[max(0, lotNumberRow[a]):min(len(CSB6400802Data), lotNumberRow[a] + 10), CSB6400802Data.columns.get_loc(lotNumberColumn[a]):CSB6400802Data.columns.get_loc(lotNumberColumn[a]) + 5]

            average1 = inspectionData.iloc[3].mean()

            minimum1 = inspectionData.iloc[3].min()

            maximum1 = inspectionData.iloc[3].max()

            csbTotalAverage1.append(average1)

            csbTotalMinimum1.append(minimum1)

            csbTotalMaximum1.append(maximum1)

            inspectionData

        csbTotalAverage1 = statistics.mean(csbTotalAverage1)

        csbTotalMinimum1 = min(csbTotalMinimum1)

        csbTotalMaximum1 = max(csbTotalMaximum1)

        csbTotalAverage1 = f"{csbTotalAverage1:.2f}"

        csbTotalMinimum1 = f"{csbTotalMinimum1:.2f}"

        csbTotalMaximum1 = f"{csbTotalMaximum1:.2f}"

        print(f"Average:{csbTotalAverage1}")

        print(f"Minimum:{csbTotalMinimum1}")

        print(f"Maximum:{csbTotalMaximum1}")
    except:
        csbTotalAverage1 = "Grade C"

        csbTotalMinimum1 = "Grade C"

        csbTotalMaximum1 = "Grade C"

        print(f"Average:{csbTotalAverage1}")

        print(f"Minimum:{csbTotalMinimum1}")

        print(f"Maximum:{csbTotalMaximum1}")

# %%
# DateAndTimeManager.GetDateToday()
# ReadDFBSnap("20241031-A")
# ReadDFB6600600()

# dfbCode
# dfbLotNumber2

# %%
def GettingDFB6600600():
    global DFB6600600Data

    global dfbCode
    global dfbLotNumber2

    global dfbTotalAverage1
    global dfbTotalAverage2
    global dfbTotalAverage3
    global dfbTotalAverage4

    global dfbTotalMinimum1
    global dfbTotalMinimum2
    global dfbTotalMinimum3
    global dfbTotalMinimum4

    global dfbTotalMaximum1
    global dfbTotalMaximum2
    global dfbTotalMaximum3
    global dfbTotalMaximum4

    dfbTotalAverage1 = []
    dfbTotalAverage2 = []
    dfbTotalAverage3 = []
    dfbTotalAverage4 = []

    dfbTotalMinimum1 = []
    dfbTotalMinimum2 = []
    dfbTotalMinimum3 = []
    dfbTotalMinimum4 = []

    dfbTotalMaximum1 = []
    dfbTotalMaximum2 = []
    dfbTotalMaximum3 = []
    dfbTotalMaximum4 = []

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    
    
    vt1Directory = (fr'\\192.168.2.19\quality control\{DateAndTimeManager.yearNow}\1.Supplier{"'"}s  Relation\A. Inspection Standard & Reference\5.) RECEIVING INSPECTION RECORD\TAKAISHI')
    
    #Finding A Folder That Contains New Trend
    for d in os.listdir(vt1Directory):
        if 'new trend' in d.lower():
            vt1Directory = os.path.join(vt1Directory, d)
            print(f"Updated vt1Directory: {vt1Directory}")
            break

    os.chdir(vt1Directory)
    
    #Finding All xlsm Files In The Current Directory
    files = glob.glob('*.xlsm')

    recentTime = 0

    #Checking Each Files In Files;
    for f in files:
        if dfbCode[:-3] in f:
            #Checking If It Is Recent File
            fileTime = os.path.getmtime(f)
            if fileTime > recentTime:
                recentTime = fileTime
                fileName = f

    # wb = load_workbook(filename=fileName, data_only=True)
    # sheet = wb['format']
    # DFB6600600Data = pd.DataFrame(sheet.values)
    # DFB6600600Data = DFB6600600Data.replace(r'\s+', '', regex=True)

    workbook = CalamineWorkbook.from_path(fileName)
    DFB6600600Data = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
    DFB6600600Data = pd.DataFrame(DFB6600600Data)
    DFB6600600Data = DFB6600600Data.replace(r'\s+', '', regex=True)

    print(fileName)
   
    #Getting The Row, Column Location Of HIBLOW
    findHiblow = [(index, column) for index, row in DFB6600600Data.iterrows() for column, value in row.items() if value == "HIBLOW"]
    hiblowRow = [index for index, _ in findHiblow]
    hiblowColumn = [column for _, column in findHiblow]

    print("Row indices:", hiblowRow)
    print("Column names:", hiblowColumn)

    # Get the Neighboring Data Of Hiblow
    hiblowFiltered = DFB6600600Data.iloc[max(0, hiblowRow[0] - 3):min(len(DFB6600600Data), hiblowRow[0] + 10), DFB6600600Data.columns.get_loc(hiblowColumn[0]):DFB6600600Data.columns.get_loc(hiblowColumn[0]) + 999]
    
    #Getting The Row, Column Location Of Lot Number
    findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == dfbLotNumber2[:-3]]
    lotNumberRow = [index for index, _ in findLotNumber]
    lotNumberColumn = [column for _, column in findLotNumber]

    print("Row indices:", lotNumberRow)
    print("Column names:", lotNumberColumn)

    # inspectionData = DFB6600600Data.iloc[max(0, lotNumberRow[0]):min(len(DFB6600600Data), lotNumberRow[0] + 10), DFB6600600Data.columns.get_loc(lotNumberColumn[0]):DFB6600600Data.columns.get_loc(lotNumberColumn[0]) + 5]

    for a in range(0, len(lotNumberColumn)):
        # Get The Neighboring Data of Lot Number
        inspectionData = DFB6600600Data.iloc[max(0, lotNumberRow[a]):min(len(DFB6600600Data), lotNumberRow[a] + 10), DFB6600600Data.columns.get_loc(lotNumberColumn[a]):DFB6600600Data.columns.get_loc(lotNumberColumn[a]) + 5]

        average1 = inspectionData.iloc[3].mean()
        average2 = inspectionData.iloc[4].mean()
        average3 = inspectionData.iloc[5].mean()
        average4 = inspectionData.iloc[6].mean()

        minimum1 = inspectionData.iloc[3].min()
        minimum2 = inspectionData.iloc[4].min()
        minimum3 = inspectionData.iloc[5].min()
        minimum4 = inspectionData.iloc[6].min()

        maximum1 = inspectionData.iloc[3].max()
        maximum2 = inspectionData.iloc[4].max()
        maximum3 = inspectionData.iloc[5].max()
        maximum4 = inspectionData.iloc[6].max()

        dfbTotalAverage1.append(average1)
        dfbTotalAverage2.append(average2)
        dfbTotalAverage3.append(average3)
        dfbTotalAverage4.append(average4)

        dfbTotalMinimum1.append(minimum1)
        dfbTotalMinimum2.append(minimum2)
        dfbTotalMinimum3.append(minimum3)
        dfbTotalMinimum4.append(minimum4)

        dfbTotalMaximum1.append(maximum1)
        dfbTotalMaximum2.append(maximum2)
        dfbTotalMaximum3.append(maximum3)
        dfbTotalMaximum4.append(maximum4)

    dfbTotalAverage1 = statistics.mean(dfbTotalAverage1)
    dfbTotalAverage2 = statistics.mean(dfbTotalAverage2)
    dfbTotalAverage3 = statistics.mean(dfbTotalAverage3)
    dfbTotalAverage4 = statistics.mean(dfbTotalAverage4)

    dfbTotalMinimum1 = min(dfbTotalMinimum1)
    dfbTotalMinimum2 = min(dfbTotalMinimum2)
    dfbTotalMinimum3 = min(dfbTotalMinimum3)
    dfbTotalMinimum4 = min(dfbTotalMinimum4)

    dfbTotalMaximum1 = max(dfbTotalMaximum1)
    dfbTotalMaximum2 = max(dfbTotalMaximum2)
    dfbTotalMaximum3 = max(dfbTotalMaximum3)
    dfbTotalMaximum4 = max(dfbTotalMaximum4)

    dfbTotalAverage1 = f"{dfbTotalAverage1:.2f}"
    dfbTotalAverage2 = f"{dfbTotalAverage2:.2f}"
    dfbTotalAverage3 = f"{dfbTotalAverage3:.2f}"
    dfbTotalAverage4 = f"{dfbTotalAverage4:.2f}"

    dfbTotalMinimum1 = f"{dfbTotalMinimum1:.2f}"
    dfbTotalMinimum2 = f"{dfbTotalMinimum2:.2f}"
    dfbTotalMinimum3 = f"{dfbTotalMinimum3:.2f}"
    dfbTotalMinimum4 = f"{dfbTotalMinimum4:.2f}"

    dfbTotalMaximum1 = f"{dfbTotalMaximum1:.2f}"
    dfbTotalMaximum2 = f"{dfbTotalMaximum2:.2f}"
    dfbTotalMaximum3 = f"{dfbTotalMaximum3:.2f}"
    dfbTotalMaximum4 = f"{dfbTotalMaximum4:.2f}"

    print(f"Average:{dfbTotalAverage1}")
    print(f"Average:{dfbTotalAverage2}")
    print(f"Average:{dfbTotalAverage3}")
    print(f"Average:{dfbTotalAverage4}")

    print(f"Minimum:{dfbTotalMinimum1}")
    print(f"Minimum:{dfbTotalMinimum2}")
    print(f"Minimum:{dfbTotalMinimum3}")
    print(f"Minimum:{dfbTotalMinimum4}")

    print(f"Maximum:{dfbTotalMaximum1}")
    print(f"Maximum:{dfbTotalMaximum2}")
    print(f"Maximum:{dfbTotalMaximum3}")
    print(f"Maximum:{dfbTotalMaximum4}")

    # return inspectionData

# %%
def GettingRDB5200200():
    global RDB5200200Data
    global rdbLotNumber2
    global rdbLotNumber3

    global rdbTotalAverage1
    global rdbTotalAverage2
    global rdbTotalAverage3
    global rdbTotalAverage4
    global rdbTotalAverage5
    global rdbTotalAverage6
    global rdbTotalAverage8

    global rdbTotalMinimum1
    global rdbTotalMinimum2
    global rdbTotalMinimum3
    global rdbTotalMinimum4
    global rdbTotalMinimum5
    global rdbTotalMinimum6
    global rdbTotalMinimum8

    global rdbTotalMaximum1
    global rdbTotalMaximum2
    global rdbTotalMaximum3
    global rdbTotalMaximum4
    global rdbTotalMaximum5
    global rdbTotalMaximum6
    global rdbTotalMaximum8

    global rdbNoDataFound

    rdbTotalAverage1 = []
    rdbTotalAverage2 = []
    rdbTotalAverage3 = []
    rdbTotalAverage4 = []
    rdbTotalAverage5 = []
    rdbTotalAverage6 = []
    rdbTotalAverage8 = []

    rdbTotalMinimum1 = []
    rdbTotalMinimum2 = []
    rdbTotalMinimum3 = []
    rdbTotalMinimum4 = []
    rdbTotalMinimum5 = []
    rdbTotalMinimum6 = []
    rdbTotalMinimum8 = []

    rdbTotalMaximum1 = []
    rdbTotalMaximum2 = []
    rdbTotalMaximum3 = []
    rdbTotalMaximum4 = []
    rdbTotalMaximum5 = []
    rdbTotalMaximum6 = []
    rdbTotalMaximum8 = []

    if not rdbNoDataFound:

        #Getting The Row, Column Location Of HIBLOW
        findHiblow = [(index, column) for index, row in RDB5200200Data.iterrows() for column, value in row.items() if value == "HIBLOW"]
        hiblowRow = [index for index, _ in findHiblow]
        hiblowColumn = [column for _, column in findHiblow]

        print("Row indices:", hiblowRow)
        print("Column names:", hiblowColumn)

        # Get the Neighboring Data Of Hiblow
        hiblowFiltered = RDB5200200Data.iloc[max(0, hiblowRow[0] - 3):min(len(RDB5200200Data), hiblowRow[0] + 10), RDB5200200Data.columns.get_loc(hiblowColumn[0]):RDB5200200Data.columns.get_loc(hiblowColumn[0]) + 999]

        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in hiblowFiltered.iterrows() for column, value in row.items() if value == rdbLotNumber3]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print("Row indices:", lotNumberRow)
        print("Column names:", lotNumberColumn)

        for a in range(0, len(lotNumberColumn)):
            # Get The Neighboring Data of Lot Number
            inspectionData = RDB5200200Data.iloc[max(0, lotNumberRow[a]):min(len(RDB5200200Data), lotNumberRow[a] + 11), RDB5200200Data.columns.get_loc(lotNumberColumn[a]):RDB5200200Data.columns.get_loc(lotNumberColumn[a]) + 5]

            average1 = inspectionData.iloc[3].mean()
            average2 = inspectionData.iloc[4].mean()
            average3 = inspectionData.iloc[5].mean()
            average4 = inspectionData.iloc[6].mean()
            average5 = inspectionData.iloc[7].mean()
            average6 = inspectionData.iloc[8].mean()
            average8 = inspectionData.iloc[10].mean()

            minimum1 = inspectionData.iloc[3].min()
            minimum2 = inspectionData.iloc[4].min()
            minimum3 = inspectionData.iloc[5].min()
            minimum4 = inspectionData.iloc[6].min()
            minimum5 = inspectionData.iloc[7].min()
            minimum6 = inspectionData.iloc[8].min()
            minimum8 = inspectionData.iloc[10].min()

            maximum1 = inspectionData.iloc[3].max()
            maximum2 = inspectionData.iloc[4].max()
            maximum3 = inspectionData.iloc[5].max()
            maximum4 = inspectionData.iloc[6].max()
            maximum5 = inspectionData.iloc[7].max()
            maximum6 = inspectionData.iloc[8].max()
            maximum8 = inspectionData.iloc[10].max()

            rdbTotalAverage1.append(average1)
            rdbTotalAverage2.append(average2)
            rdbTotalAverage3.append(average3)
            rdbTotalAverage4.append(average4)
            rdbTotalAverage5.append(average5)
            rdbTotalAverage6.append(average6)
            rdbTotalAverage8.append(average8)

            rdbTotalMinimum1.append(minimum1)
            rdbTotalMinimum2.append(minimum2)
            rdbTotalMinimum3.append(minimum3)
            rdbTotalMinimum4.append(minimum4)
            rdbTotalMinimum5.append(minimum5)
            rdbTotalMinimum6.append(minimum6)
            rdbTotalMinimum8.append(minimum8)

            rdbTotalMaximum1.append(maximum1)
            rdbTotalMaximum2.append(maximum2)
            rdbTotalMaximum3.append(maximum3)
            rdbTotalMaximum4.append(maximum4)
            rdbTotalMaximum5.append(maximum5)
            rdbTotalMaximum6.append(maximum6)
            rdbTotalMaximum8.append(maximum8)

        rdbTotalAverage1 = statistics.mean(rdbTotalAverage1)
        rdbTotalAverage2 = statistics.mean(rdbTotalAverage2)
        rdbTotalAverage3 = statistics.mean(rdbTotalAverage3)
        rdbTotalAverage4 = statistics.mean(rdbTotalAverage4)
        rdbTotalAverage5 = statistics.mean(rdbTotalAverage5)
        rdbTotalAverage6 = statistics.mean(rdbTotalAverage6)
        rdbTotalAverage8 = statistics.mean(rdbTotalAverage8)

        rdbTotalMinimum1 = min(rdbTotalMinimum1)
        rdbTotalMinimum2 = min(rdbTotalMinimum2)
        rdbTotalMinimum3 = min(rdbTotalMinimum3)
        rdbTotalMinimum4 = min(rdbTotalMinimum4)
        rdbTotalMinimum5 = min(rdbTotalMinimum5)
        rdbTotalMinimum6 = min(rdbTotalMinimum6)
        rdbTotalMinimum8 = min(rdbTotalMinimum8)

        rdbTotalMaximum1 = max(rdbTotalMaximum1)
        rdbTotalMaximum2 = max(rdbTotalMaximum2)
        rdbTotalMaximum3 = max(rdbTotalMaximum3)
        rdbTotalMaximum4 = max(rdbTotalMaximum4)
        rdbTotalMaximum5 = max(rdbTotalMaximum5)
        rdbTotalMaximum6 = max(rdbTotalMaximum6)
        rdbTotalMaximum8 = max(rdbTotalMaximum8)

        rdbTotalAverage1 = f"{rdbTotalAverage1:.2f}"
        rdbTotalAverage2 = f"{rdbTotalAverage2:.2f}"
        rdbTotalAverage3 = f"{rdbTotalAverage3:.2f}"
        rdbTotalAverage4 = f"{rdbTotalAverage4:.2f}"
        rdbTotalAverage5 = f"{rdbTotalAverage5:.2f}"
        rdbTotalAverage6 = f"{rdbTotalAverage6:.2f}"
        rdbTotalAverage8 = f"{rdbTotalAverage8:.2f}"

        rdbTotalMinimum1 = f"{rdbTotalMinimum1:.2f}"
        rdbTotalMinimum2 = f"{rdbTotalMinimum2:.2f}"
        rdbTotalMinimum3 = f"{rdbTotalMinimum3:.2f}"
        rdbTotalMinimum4 = f"{rdbTotalMinimum4:.2f}"
        rdbTotalMinimum5 = f"{rdbTotalMinimum5:.2f}"
        rdbTotalMinimum6 = f"{rdbTotalMinimum6:.2f}"
        rdbTotalMinimum8 = f"{rdbTotalMinimum8:.2f}"

        rdbTotalMaximum1 = f"{rdbTotalMaximum1:.2f}"
        rdbTotalMaximum2 = f"{rdbTotalMaximum2:.2f}"
        rdbTotalMaximum3 = f"{rdbTotalMaximum3:.2f}"
        rdbTotalMaximum4 = f"{rdbTotalMaximum4:.2f}"
        rdbTotalMaximum5 = f"{rdbTotalMaximum5:.2f}"
        rdbTotalMaximum6 = f"{rdbTotalMaximum6:.2f}"
        rdbTotalMaximum8 = f"{rdbTotalMaximum8:.2f}"

        print(f"RDB Total Average 1: {rdbTotalAverage1}")
        print(f"RDB Total Average 2: {rdbTotalAverage2}")
        print(f"RDB Total Average 3: {rdbTotalAverage3}")
        print(f"RDB Total Average 4: {rdbTotalAverage4}")
        print(f"RDB Total Average 5: {rdbTotalAverage5}")
        print(f"RDB Total Average 6: {rdbTotalAverage6}")
        print(f"RDB Total Average 8: {rdbTotalAverage8}")

        print(f"RDB Total Minimum 1: {rdbTotalMinimum1}")
        print(f"RDB Total Minimum 2: {rdbTotalMinimum2}")
        print(f"RDB Total Minimum 3: {rdbTotalMinimum3}")
        print(f"RDB Total Minimum 4: {rdbTotalMinimum4}")
        print(f"RDB Total Minimum 5: {rdbTotalMinimum5}")
        print(f"RDB Total Minimum 6: {rdbTotalMinimum6}")
        print(f"RDB Total Minimum 8: {rdbTotalMinimum8}")

        print(f"RDB Total Maximum 1: {rdbTotalMaximum1}")
        print(f"RDB Total Maximum 2: {rdbTotalMaximum2}")
        print(f"RDB Total Maximum 3: {rdbTotalMaximum3}")
        print(f"RDB Total Maximum 4: {rdbTotalMaximum4}")
        print(f"RDB Total Maximum 5: {rdbTotalMaximum5}")
        print(f"RDB Total Maximum 6: {rdbTotalMaximum6}")
        print(f"RDB Total Maximum 8: {rdbTotalMaximum8}")

# %%
# DateAndTimeManager.GetDateToday()
# ReadRDB5200200CheckSheet("20241021-F")
# ReadRDB5200200()
# GettingRDB5200200()

# %%
# DateAndTimeManager.GetDateToday()
# ReadRDB5200200CheckSheet("20241018-F")
# ReadRDB5200200()
# GettingRDB5200200()

# %%
# rdbTotalAverage1

# %%
# DateAndTimeManager.GetDateToday()
# ReadDFBSnap("20241031-A")
# ReadDFB6600600()
# dfbCode
# GettingDFB6600600()



# %%
from Imports import *
import DateAndTimeManager

# %%
class em2P():
    em2PData = ""
    em2PItemCode = ""

    totalAverage3 = []
    totalAverage4 = []
    totalAverage5 = []
    totalAverage10 = []

    totalMinimum3 = []
    totalMinimum4 = []
    totalMinimum5 = []
    
    totalMaximum3 = []
    totalMaximum4 = []
    totalMaximum5 = []

    readingYear = ""
    fileFinishedReading = False
    fileList = []

    isValueRetrieve = False

    def __init__(self):
        pass
    def ReadExcel(self, itemCode):
        self.em2PItemCode = itemCode
        self.fileList = []

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        while not self.fileFinishedReading:
            try:
                vt1Directory = (fr'\\192.168.2.19\quality control\{str(self.readingYear)}')

                for d in os.listdir(vt1Directory):
                    if "supplier" in d.lower():
                        vt1Directory = os.path.join(vt1Directory, d)
                        for d in os.listdir(vt1Directory):
                            if "inspection standard" in d.lower():
                                vt1Directory = os.path.join(vt1Directory, d)
                                for d in os.listdir(vt1Directory):
                                    if "receiving inspection record" in d.lower():
                                        vt1Directory = os.path.join(vt1Directory, d)
                                        
                                        #CHECKING THE ITEM CODE
                                        if itemCode == "EM0580106P":
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "gaptec" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0580106P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em2PData = pd.DataFrame(self.em2PData)
                                                            self.em2PData = self.em2PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em2PData)
                                            except:
                                                print("NO DATA FOUND IN GAPTEC")
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "dhye" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0580106P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em2PData = pd.DataFrame(self.em2PData)
                                                            self.em2PData = self.em2PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em2PData)
                                            except:
                                                print("NO DATA FOUND IN DHYE")

                                        elif itemCode == "EM0660046P":
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "gaptec" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660046P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em2PData = pd.DataFrame(self.em2PData)
                                                            self.em2PData = self.em2PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em2PData)
                                            except:
                                                print("NO DATA FOUND IN GAPTEC")

                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "dhye" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660046P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em2PData = pd.DataFrame(self.em2PData)
                                                            self.em2PData = self.em2PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em2PData)
                                            except:
                                                print("NO DATA FOUND IN DHYE")

                                        elif itemCode == "EM0660044P":
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "gaptec" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660044P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em2PData = pd.DataFrame(self.em2PData)
                                                            self.em2PData = self.em2PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em2PData)
                                            except:
                                                print("NO DATA FOUND IN GAPTEC")

                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "dhye" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660046P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em2PData = pd.DataFrame(self.em2PData)
                                                            self.em2PData = self.em2PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em2PData)
                                            except:
                                                print("NO DATA FOUND IN DHYE")
            except:                             
                pass
            
            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                self.fileFinishedReading = True

        # self.em2PData.replace('', np.nan, inplace=True)

        for file in self.fileList:
            file.replace('', np.nan, inplace=True)   
                
    def GettingData(self, lotNumber):
        for fileNum in range(len(self.fileList)):
            self.totalAverage3 = []
            self.totalAverage4 = []
            self.totalAverage5 = []
            self.totalAverage10 = []

            self.totalMinimum3 = []
            self.totalMinimum4 = []
            self.totalMinimum5 = []
            
            self.totalMaximum3 = []
            self.totalMaximum4 = []
            self.totalMaximum5 = []
            
            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of SUPPLIER
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of SUPPLIER
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
            supplierFiltered

            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    # Get The Neighboring Data of Lot Number
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    #CHECKING THE ITEM CODE
                    if self.em2PItemCode == "EM0580106P":
                        average3 = inspectionData.iloc[5].mean()
                        average4 = inspectionData.iloc[6].mean()
                        average5 = inspectionData.iloc[7].mean()
                        average10 = inspectionData.iloc[12, 0]

                        minimum3 = inspectionData.iloc[5].min()
                        minimum4 = inspectionData.iloc[6].min()
                        minimum5 = inspectionData.iloc[7].min()

                        maximum3 = inspectionData.iloc[5].max()
                        maximum4 = inspectionData.iloc[6].max()
                        maximum5 = inspectionData.iloc[7].max()

                        self.totalAverage3.append(average3)
                        self.totalAverage4.append(average4)
                        self.totalAverage5.append(average5)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)
                        self.totalMinimum4.append(minimum4)
                        self.totalMinimum5.append(minimum5)

                        self.totalMaximum3.append(maximum3)
                        self.totalMaximum4.append(maximum4)
                        self.totalMaximum5.append(maximum5)
                        
                    elif self.em2PItemCode == "EM0660046P":
                        average3 = inspectionData.iloc[5].mean()
                        average10 = inspectionData.iloc[8, 0]

                        minimum3 = inspectionData.iloc[5].min()

                        maximum3 = inspectionData.iloc[5].max()

                        self.totalAverage3.append(average3)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)

                        self.totalMaximum3.append(maximum3)

                    elif self.em2PItemCode == "EM0660044P":
                        average3 = inspectionData.iloc[5].mean()
                        average10 = inspectionData.iloc[8, 0]

                        minimum3 = inspectionData.iloc[5].min()

                        maximum3 = inspectionData.iloc[5].max()

                        self.totalAverage3.append(average3)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)

                        self.totalMaximum3.append(maximum3)

                #CHECKING THE ITEM CODE
                if self.em2PItemCode == "EM0580106P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage4 = statistics.mean(self.totalAverage4)
                    self.totalAverage5 = statistics.mean(self.totalAverage5)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)
                    self.totalMinimum4 = min(self.totalMinimum4)
                    self.totalMinimum5 = min(self.totalMinimum5)

                    self.totalMaximum3 = max(self.totalMaximum3)
                    self.totalMaximum4 = max(self.totalMaximum4)
                    self.totalMaximum5 = max(self.totalMaximum5)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage4 = f"{self.totalAverage4:.2f}"
                    self.totalAverage5 = f"{self.totalAverage5:.2f}"
                    self.totalAverage10 = f"{self.totalAverage10:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"
                    self.totalMinimum4 = f"{self.totalMinimum4:.2f}"
                    self.totalMinimum5 = f"{self.totalMinimum5:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    self.totalMaximum4 = f"{self.totalMaximum4:.2f}"
                    self.totalMaximum5 = f"{self.totalMaximum5:.2f}"

                    break
                elif self.em2PItemCode == "EM0660046P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)

                    self.totalMaximum3 = max(self.totalMaximum3)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"

                    break
                elif self.em2PItemCode == "EM0660044P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)

                    self.totalMaximum3 = max(self.totalMaximum3)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"

                    break
            except:
                self.totalAverage3 = "No Data Found"
                self.totalAverage4 = "No Data Found"
                self.totalAverage5 = "No Data Found"
                self.totalAverage10 = "No Data Found"

                self.totalMinimum3 = "No Data Found"
                self.totalMinimum4 = "No Data Found"
                self.totalMinimum5 = "No Data Found"

                self.totalMaximum3 = "No Data Found"
                self.totalMaximum4 = "No Data Found"
                self.totalMaximum5 = "No Data Found"

            print(f"Total Average: {self.totalAverage3}")
            print(f"Total Average: {self.totalAverage4}")
            print(f"Total Average: {self.totalAverage5}")
            print(f"Total Average: {self.totalAverage10}")

            print(f"Total Minimum: {self.totalMinimum3}")
            print(f"Total Minimum: {self.totalMinimum4}")
            print(f"Total Minimum: {self.totalMinimum5}")

            print(f"Total Maximum: {self.totalMaximum3}")
            print(f"Total Maximum: {self.totalMaximum4}")
            print(f"Total Maximum: {self.totalMaximum5}")

        print(f"Selected Total Average: {self.totalAverage3}")
        print(f"Selected Total Average: {self.totalAverage4}")
        print(f"Selected Total Average: {self.totalAverage5}")
        print(f"Selected Total Average: {self.totalAverage10}")

        print(f"Selected Total Minimum: {self.totalMinimum3}")
        print(f"Selected Total Minimum: {self.totalMinimum4}")
        print(f"Selected Total Minimum: {self.totalMinimum5}")

        print(f"Selected Total Maximum: {self.totalMaximum3}")
        print(f"Selected Total Maximum: {self.totalMaximum4}")
        print(f"Selected Total Maximum: {self.totalMaximum5}")

    def Trial(self, lotNumber):
        fileNum = 1

        #Getting The Row, Column Location Of SUPPLIER
        findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
        supplierRow = [index for index, _ in findSupplier]
        supplierColumn = [column for _, column in findSupplier]

        print("Row indices:", supplierRow)
        print("Column names:", supplierColumn)

        # Get the Neighboring Data Of SUPPLIER
        supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
        supplierFiltered

        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print("Row indices:", lotNumberRow)
        print("Column names:", lotNumberColumn)

        inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[3]):min(len(self.fileList[fileNum]), lotNumberRow[3] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[3]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[3]) + 5]

        return inspectionData

# %%
# em2p = em2P()
# DateAndTimeManager.GetDateToday()
# em2p.readingYear = int(DateAndTimeManager.yearNow)

# em2p.ReadExcel("EM0580106P")
# em2p.ReadExcel("EM0580106P")
# # em2p.ReadExcel("EM0660046P")
# em2p.ReadExcel("EM0660044P")

# print(f"Total Number Of Files {len(em2p.fileList)}")

# em2p.GettingData("CAT-4J15DI")
# em2p.GettingData("FC6030-3E04GT")
# # em2p.GettingData("CAT-5A07DI")
# # em2p.GettingData("CAT-5A06DI")
# em2p.GettingData("FC6030-4G26GT")
# # em2p.GettingData("FC6030-4F05GT")

# # em2p.Trial("FC6030-3E04GT")
# # em2p.Trial("FC6030-4F05GT")

# %%
class em3P():
    em3PData = ""
    em3PItemCode = ""

    totalAverage3 = []
    totalAverage4 = []
    totalAverage5 = []
    totalAverage10 = []

    totalMinimum3 = []
    totalMinimum4 = []
    totalMinimum5 = []
    
    totalMaximum3 = []
    totalMaximum4 = []
    totalMaximum5 = []

    readingYear = ""
    fileFinishedReading = False
    fileList = []

    isValueRetrieve = False

    def __init__(self):
        pass
    def ReadExcel(self, itemCode):
        self.em3PItemCode = itemCode
        self.fileList = []

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        while not self.fileFinishedReading:
            try:
                vt1Directory = (fr'\\192.168.2.19\quality control\{str(self.readingYear)}')

                for d in os.listdir(vt1Directory):
                    if "supplier" in d.lower():
                        vt1Directory = os.path.join(vt1Directory, d)
                        for d in os.listdir(vt1Directory):
                            if "inspection standard" in d.lower():
                                vt1Directory = os.path.join(vt1Directory, d)
                                for d in os.listdir(vt1Directory):
                                    if "receiving inspection record" in d.lower():
                                        vt1Directory = os.path.join(vt1Directory, d)
                                        
                                        #CHECKING THE ITEM CODE
                                        if itemCode == "EM0580107P":
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "gaptec" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0580107P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em3PData = pd.DataFrame(self.em3PData)
                                                            self.em3PData = self.em3PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em3PData)
                                            except:
                                                print("NO DATA FOUND IN GAPTEC")
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "dhye" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0580107P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em3PData = pd.DataFrame(self.em3PData)
                                                            self.em3PData = self.em3PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em3PData)
                                            except:
                                                print("NO DATA FOUND IN DHYE")
                                        elif itemCode == "EM0660047P":
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "gaptec" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660047P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em3PData = pd.DataFrame(self.em3PData)
                                                            self.em3PData = self.em3PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em3PData)
                                            except:
                                                print("NO DATA FOUND IN GAPTEC")
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "dhye" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660047P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em3PData = pd.DataFrame(self.em3PData)
                                                            self.em3PData = self.em3PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em3PData)
                                            except:
                                                print("NO DATA FOUND IN DHYE")
                                        elif itemCode == "EM0660045P":
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "gaptec" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660045P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em3PData = pd.DataFrame(self.em3PData)
                                                            self.em3PData = self.em3PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em3PData)
                                            except:
                                                print("NO DATA FOUND IN GAPTEC")
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "dhye" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*EM0660045P*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.em3PData = pd.DataFrame(self.em3PData)
                                                            self.em3PData = self.em3PData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.em3PData)
                                            except:
                                                print("NO DATA FOUND IN DHYE")
            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                self.fileFinishedReading = True                  
                                
        # self.em3PData.replace('', np.nan, inplace=True)  

        for file in self.fileList:
            file.replace('', np.nan, inplace=True)            
                                
    def GettingData(self, lotNumber):
        for fileNum in range(len(self.fileList)):
            self.totalAverage3 = []
            self.totalAverage4 = []
            self.totalAverage5 = []
            self.totalAverage10 = []

            self.totalMinimum3 = []
            self.totalMinimum4 = []
            self.totalMinimum5 = []
            
            self.totalMaximum3 = []
            self.totalMaximum4 = []
            self.totalMaximum5 = []
            
            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of HIBLOW
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of SUPPLIER
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
            supplierFiltered

            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    # Get The Neighboring Data of Lot Number
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    #CHECKING THE ITEM CODE
                    if self.em3PItemCode == "EM0580107P":
                        average3 = inspectionData.iloc[5].mean()
                        average4 = inspectionData.iloc[6].mean()
                        average5 = inspectionData.iloc[7].mean()
                        average10 = inspectionData.iloc[12, 0]

                        minimum3 = inspectionData.iloc[5].min()
                        minimum4 = inspectionData.iloc[6].min()
                        minimum5 = inspectionData.iloc[7].min()

                        maximum3 = inspectionData.iloc[5].max()
                        maximum4 = inspectionData.iloc[6].max()
                        maximum5 = inspectionData.iloc[7].max()

                        self.totalAverage3.append(average3)
                        self.totalAverage4.append(average4)
                        self.totalAverage5.append(average5)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)
                        self.totalMinimum4.append(minimum4)
                        self.totalMinimum5.append(minimum5)

                        self.totalMaximum3.append(maximum3)
                        self.totalMaximum4.append(maximum4)
                        self.totalMaximum5.append(maximum5)
                    #CHECKING THE ITEM CODE
                    elif self.em3PItemCode == "EM0660047P":
                        average3 = inspectionData.iloc[5].mean()
                        average10 = inspectionData.iloc[8, 0]

                        minimum3 = inspectionData.iloc[5].min()
                    
                        maximum3 = inspectionData.iloc[5].max()
                        
                        self.totalAverage3.append(average3)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)

                        self.totalMaximum3.append(maximum3)
                    #CHECKING THE ITEM CODE
                    elif self.em3PItemCode == "EM0660045P":
                        average3 = inspectionData.iloc[5].mean()
                        average10 = inspectionData.iloc[8, 0]

                        minimum3 = inspectionData.iloc[5].min()
                    
                        maximum3 = inspectionData.iloc[5].max()
                        
                        self.totalAverage3.append(average3)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)

                        self.totalMaximum3.append(maximum3)
                    
                #CHECKING THE ITEM CODE
                if self.em3PItemCode == "EM0580107P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage4 = statistics.mean(self.totalAverage4)
                    self.totalAverage5 = statistics.mean(self.totalAverage5)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)
                    self.totalMinimum4 = min(self.totalMinimum4)
                    self.totalMinimum5 = min(self.totalMinimum5)

                    self.totalMaximum3 = max(self.totalMaximum3)
                    self.totalMaximum4 = max(self.totalMaximum4)
                    self.totalMaximum5 = max(self.totalMaximum5)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage4 = f"{self.totalAverage4:.2f}"
                    self.totalAverage5 = f"{self.totalAverage5:.2f}"
                    self.totalAverage10 = f"{self.totalAverage10:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"
                    self.totalMinimum4 = f"{self.totalMinimum4:.2f}"
                    self.totalMinimum5 = f"{self.totalMinimum5:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    self.totalMaximum4 = f"{self.totalMaximum4:.2f}"
                    self.totalMaximum5 = f"{self.totalMaximum5:.2f}"

                    break
                #CHECKING THE ITEM CODE
                elif self.em3PItemCode == "EM0660047P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)

                    self.totalMaximum3 = max(self.totalMaximum3)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage10 = f"{self.totalAverage10:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    break
                #CHECKING THE ITEM CODE
                elif self.em3PItemCode == "EM0660045P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)

                    self.totalMaximum3 = max(self.totalMaximum3)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage10 = f"{self.totalAverage10:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    break
            except:
                self.totalAverage3 = "No Data Found"
                self.totalAverage4 = "No Data Found"
                self.totalAverage5 = "No Data Found"
                self.totalAverage10 = "No Data Found"

                self.totalMinimum3 = "No Data Found"
                self.totalMinimum4 = "No Data Found"
                self.totalMinimum5 = "No Data Found"

                self.totalMaximum3 = "No Data Found"
                self.totalMaximum4 = "No Data Found"
                self.totalMaximum5 = "No Data Found"

            print(f"Total Average: {self.totalAverage3}")
            print(f"Total Average: {self.totalAverage4}")
            print(f"Total Average: {self.totalAverage5}")
            print(f"Total Average: {self.totalAverage10}")

            print(f"Total Minimum: {self.totalMinimum3}")
            print(f"Total Minimum: {self.totalMinimum4}")
            print(f"Total Minimum: {self.totalMinimum5}")

            print(f"Total Maximum: {self.totalMaximum3}")
            print(f"Total Maximum: {self.totalMaximum4}")
            print(f"Total Maximum: {self.totalMaximum5}")

        print(f"Selected Total Average: {self.totalAverage3}")
        print(f"Selected Total Average: {self.totalAverage4}")
        print(f"Selected Total Average: {self.totalAverage5}")
        print(f"Selected Total Average: {self.totalAverage10}")

        print(f"Selected Total Minimum: {self.totalMinimum3}")
        print(f"Selected Total Minimum: {self.totalMinimum4}")
        print(f"Selected Total Minimum: {self.totalMinimum5}")

        print(f"Selected Total Maximum: {self.totalMaximum3}")
        print(f"Selected Total Maximum: {self.totalMaximum4}")
        print(f"Selected Total Maximum: {self.totalMaximum5}")

    def Trial(self, lotNumber):
        fileNum = 0
        a = 2

        #Getting The Row, Column Location Of HIBLOW
        findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
        supplierRow = [index for index, _ in findSupplier]
        supplierColumn = [column for _, column in findSupplier]

        print("Row indices:", supplierRow)
        print("Column names:", supplierColumn)

        # Get the Neighboring Data Of SUPPLIER
        supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
        
        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print("Row indices:", lotNumberRow)
        print("Column names:", lotNumberColumn)

        inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]
        return inspectionData

# %%
# em3p = em3P()
# DateAndTimeManager.GetDateToday()
# em3p.readingYear = int(DateAndTimeManager.yearNow)
# em3p.ReadExcel("EM0580107P")
# # em3p.ReadExcel("EM0660047P")
# # em3p.ReadExcel("EM0660045P")
# print(f"Total Number Of Files {len(em3p.fileList)}")
# # em3p.GettingData("CAT-5A07DI")
# # em3p.GettingData("FC6030-4F05GT")
# # em3p.GettingData("FC6030-4G26GT")
# # em3p.Trial("CAT-5A07DI")

# %%
class fM():
    fmData = ""
    fmItemCode = ""

    totalAverage1 = []
    totalAverage2 = []
    totalAverage3 = []
    totalAverage4 = []
    totalAverage5 = []
    totalAverage6 = []
    totalAverage7 = []

    totalMinimum1 = []
    totalMinimum2 = []
    totalMinimum3 = []
    totalMinimum4 = []
    totalMinimum5 = []
    totalMinimum6 = []
    totalMinimum7 = []
    
    totalMaximum1 = []
    totalMaximum2 = []
    totalMaximum3 = []
    totalMaximum4 = []
    totalMaximum5 = []
    totalMaximum6 = []
    totalMaximum7 = []

    readingYear = ""
    fileFinishedReading = False
    fileList = []

    isValueRetrieve = False

    def __init__(self):
        pass
    def ReadExcel(self, itemCode):
        self.fmItemCode = itemCode
        self.fileList = []

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        while not self.fileFinishedReading:
            try:
                vt1Directory = (fr'\\192.168.2.19\quality control\{str(self.readingYear)}')

                for d in os.listdir(vt1Directory):
                    if "supplier" in d.lower():
                        vt1Directory = os.path.join(vt1Directory, d)
                        for d in os.listdir(vt1Directory):
                            if "inspection standard" in d.lower():
                                vt1Directory = os.path.join(vt1Directory, d)
                                for d in os.listdir(vt1Directory):
                                    if "receiving inspection record" in d.lower():
                                        vt1Directory = os.path.join(vt1Directory, d)

                                        #CHECKING THE ITEM CODE
                                        if itemCode == "FM05000102-00A" or itemCode == "FM05000102-01A":
                                            try:
                                                for d in os.listdir(vt1Directory):
                                                    if "cronics" in d.lower():
                                                        directory = os.path.join(vt1Directory, d)

                                                        #Finding A Folder That Contains New Trend
                                                        for d in os.listdir(directory):
                                                            if 'new trend' in d.lower():
                                                                directory = os.path.join(directory, d)
                                                                print(f"Updated vt1Directory: {directory}")
                                                                break

                                                        os.chdir(directory)

                                                        files = glob.glob('*FM05000102*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.fmData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.fmData = pd.DataFrame(self.fmData)
                                                            self.fmData = self.fmData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"FM FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.fmData)
                                            except:
                                                print("NO DATA FOUND IN CRONICS")

                                        elif itemCode == "FM03500100-01":
                                            # IN PROGRESS
                                            pass
            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                self.fileFinishedReading = True   

    def GettingData(self, lotNumber):
        for fileNum in range(len(self.fileList)):
            self.totalAverage1 = []
            self.totalAverage2 = []
            self.totalAverage3 = []
            self.totalAverage4 = []
            self.totalAverage5 = []
            self.totalAverage6 = []
            self.totalAverage7 = []

            self.totalMinimum1 = []
            self.totalMinimum2 = []
            self.totalMinimum3 = []
            self.totalMinimum4 = []
            self.totalMinimum5 = []
            self.totalMinimum6 = []
            self.totalMinimum7 = []

            self.totalMaximum1 = []
            self.totalMaximum2 = []
            self.totalMaximum3 = []
            self.totalMaximum4 = []
            self.totalMaximum5 = []
            self.totalMaximum6 = []
            self.totalMaximum7 = []

            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of HIBLOW
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of SUPPLIER
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
            # return supplierFiltered

            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    # Get The Neighboring Data of Lot Number
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    #CHECKING THE ITEM CODE
                    if self.fmItemCode == "FM05000102-00A" or self.fmItemCode == "FM05000102-01A":
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

                        self.totalAverage1.append(average1)
                        self.totalAverage2.append(average2)
                        self.totalAverage3.append(average3)
                        self.totalAverage4.append(average4)
                        self.totalAverage5.append(average5)
                        self.totalAverage6.append(average6)
                        self.totalAverage7.append(average7)

                        self.totalMinimum1.append(minimum1)
                        self.totalMinimum2.append(minimum2)
                        self.totalMinimum3.append(minimum3)
                        self.totalMinimum4.append(minimum4)
                        self.totalMinimum5.append(minimum5)
                        self.totalMinimum6.append(minimum6)
                        self.totalMinimum7.append(minimum7)

                        self.totalMaximum1.append(maximum1)
                        self.totalMaximum2.append(maximum2)
                        self.totalMaximum3.append(maximum3)
                        self.totalMaximum4.append(maximum4)
                        self.totalMaximum5.append(maximum5)
                        self.totalMaximum6.append(maximum6)
                        self.totalMaximum7.append(maximum7)

                    elif self.fmItemCode == "FM03500100-01":
                        # IN PROGRESS
                        pass

                if self.fmItemCode == "FM05000102-00A" or self.fmItemCode == "FM05000102-01A":
                    self.totalAverage1 = statistics.mean(self.totalAverage1)
                    self.totalAverage2 = statistics.mean(self.totalAverage2)
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage4 = statistics.mean(self.totalAverage4)
                    self.totalAverage5 = statistics.mean(self.totalAverage5)
                    self.totalAverage6 = statistics.mean(self.totalAverage6)
                    self.totalAverage7 = statistics.mean(self.totalAverage7)

                    self.totalMinimum1 = min(self.totalMinimum1)
                    self.totalMinimum2 = min(self.totalMinimum2)
                    self.totalMinimum3 = min(self.totalMinimum3)
                    self.totalMinimum4 = min(self.totalMinimum4)
                    self.totalMinimum5 = min(self.totalMinimum5)
                    self.totalMinimum6 = min(self.totalMinimum6)
                    self.totalMinimum7 = min(self.totalMinimum7)

                    self.totalMaximum1 = max(self.totalMaximum1)
                    self.totalMaximum2 = max(self.totalMaximum2)
                    self.totalMaximum3 = max(self.totalMaximum3)
                    self.totalMaximum4 = max(self.totalMaximum4)
                    self.totalMaximum5 = max(self.totalMaximum5)
                    self.totalMaximum6 = max(self.totalMaximum6)
                    self.totalMaximum7 = max(self.totalMaximum7)

                    self.totalAverage1 = f"{self.totalAverage1:.2f}"
                    self.totalAverage2 = f"{self.totalAverage2:.2f}"
                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage4 = f"{self.totalAverage4:.2f}"
                    self.totalAverage5 = f"{self.totalAverage5:.2f}"
                    self.totalAverage6 = f"{self.totalAverage6:.2f}"
                    self.totalAverage7 = f"{self.totalAverage7:.2f}"
                    
                    self.totalMinimum1 = f"{self.totalMinimum1:.2f}"
                    self.totalMinimum2 = f"{self.totalMinimum2:.2f}"
                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"
                    self.totalMinimum4 = f"{self.totalMinimum4:.2f}"
                    self.totalMinimum5 = f"{self.totalMinimum5:.2f}"
                    self.totalMinimum6 = f"{self.totalMinimum6:.2f}"
                    self.totalMinimum7 = f"{self.totalMinimum7:.2f}"

                    self.totalMaximum1 = f"{self.totalMaximum1:.2f}"
                    self.totalMaximum2 = f"{self.totalMaximum2:.2f}"
                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    self.totalMaximum4 = f"{self.totalMaximum4:.2f}"
                    self.totalMaximum5 = f"{self.totalMaximum5:.2f}"
                    self.totalMaximum6 = f"{self.totalMaximum6:.2f}"
                    self.totalMaximum7 = f"{self.totalMaximum7:.2f}"

                    break

                elif self.fmItemCode == "FM03500100-01":
                    # IN PROGRESS
                    pass
                
            except:
                self.totalAverage1 = "No Data Found"
                self.totalAverage2 = "No Data Found"
                self.totalAverage3 = "No Data Found"
                self.totalAverage4 = "No Data Found"
                self.totalAverage5 = "No Data Found"
                self.totalAverage6 = "No Data Found"
                self.totalAverage7 = "No Data Found"

                self.totalMinimum1 = "No Data Found"
                self.totalMinimum2 = "No Data Found"
                self.totalMinimum3 = "No Data Found"
                self.totalMinimum4 = "No Data Found"
                self.totalMinimum5 = "No Data Found"
                self.totalMinimum6 = "No Data Found"
                self.totalMinimum7 = "No Data Found"

                self.totalMaximum1 = "No Data Found"
                self.totalMaximum2 = "No Data Found"
                self.totalMaximum3 = "No Data Found"
                self.totalMaximum4 = "No Data Found"
                self.totalMaximum5 = "No Data Found"
                self.totalMaximum6 = "No Data Found"
                self.totalMaximum7 = "No Data Found"

            print(f"Total Average: {self.totalAverage1}")
            print(f"Total Average: {self.totalAverage2}")
            print(f"Total Average: {self.totalAverage3}")
            print(f"Total Average: {self.totalAverage4}")
            print(f"Total Average: {self.totalAverage5}")
            print(f"Total Average: {self.totalAverage6}")
            print(f"Total Average: {self.totalAverage7}")

            print(f"Total Minimum: {self.totalMinimum1}")
            print(f"Total Minimum: {self.totalMinimum2}")
            print(f"Total Minimum: {self.totalMinimum3}")
            print(f"Total Minimum: {self.totalMinimum4}")
            print(f"Total Minimum: {self.totalMinimum5}")
            print(f"Total Minimum: {self.totalMinimum6}")
            print(f"Total Minimum: {self.totalMinimum7}")

            print(f"Total Maximum: {self.totalMaximum1}")
            print(f"Total Maximum: {self.totalMaximum2}")
            print(f"Total Maximum: {self.totalMaximum3}")
            print(f"Total Maximum: {self.totalMaximum4}")
            print(f"Total Maximum: {self.totalMaximum5}")
            print(f"Total Maximum: {self.totalMaximum6}")
            print(f"Total Maximum: {self.totalMaximum7}")

        print(f"Selected Total Average: {self.totalAverage1}")
        print(f"Selected Total Average: {self.totalAverage2}")
        print(f"Selected Total Average: {self.totalAverage3}")
        print(f"Selected Total Average: {self.totalAverage4}")
        print(f"Selected Total Average: {self.totalAverage5}")
        print(f"Selected Total Average: {self.totalAverage6}")
        print(f"Selected Total Average: {self.totalAverage7}")

        print(f"Selected Total Minimum: {self.totalMinimum1}")
        print(f"Selected Total Minimum: {self.totalMinimum2}")
        print(f"Selected Total Minimum: {self.totalMinimum3}")
        print(f"Selected Total Minimum: {self.totalMinimum4}")
        print(f"Selected Total Minimum: {self.totalMinimum5}")
        print(f"Selected Total Minimum: {self.totalMinimum6}")
        print(f"Selected Total Minimum: {self.totalMinimum7}")

        print(f"Selected Total Maximum: {self.totalMaximum1}")
        print(f"Selected Total Maximum: {self.totalMaximum2}")
        print(f"Selected Total Maximum: {self.totalMaximum3}")
        print(f"Selected Total Maximum: {self.totalMaximum4}")
        print(f"Selected Total Maximum: {self.totalMaximum5}")
        print(f"Selected Total Maximum: {self.totalMaximum6}")
        print(f"Selected Total Maximum: {self.totalMaximum7}")

# %%
# fm = fM()
# DateAndTimeManager.GetDateToday()
# fm.readingYear = int(DateAndTimeManager.yearNow)
# fm.ReadExcel("FM05000102-00A")
# # fm.ReadExcel("FM05000102-01A")
# # print(f"Total Number Of Files {len(fm.fileList)}")
# fm.GettingData("112524A-40")

# %%
class dFB():
    dfbSnapData = ""
    dfbLetterCode = ""
    dfbLotNumber = ""
    dfbMonth = ""
    dfbCode = ""
    dfbLotNumber2 = ""
    dfbYear = ""

    totalAverage1 = []
    totalAverage2 = []
    totalAverage3 = []
    totalAverage4 = []

    totalMinimum1 = []
    totalMinimum2 = []
    totalMinimum3 = []
    totalMinimum4 = []

    totalMaximum1 = []
    totalMaximum2 = []
    totalMaximum3 = []
    totalMaximum4 = []

    readingYear = ""

    fileList = []
    fileFinishedReading = False

    def __init__(self):
        pass
    def ReadDfbSnap(self, lotNumber):
        self.dfbLetterCode = lotNumber[-1]
        self.dfbYear = lotNumber[:-6]

        #Removing The Last Two Values Of Lot Number
        lotNumber = lotNumber[:-2]
        #Changing The Format Of Lot Number
        lotNumber = datetime2.strptime(lotNumber, "%Y%m%d")
        self.dfbMonth = lotNumber.strftime("%B")
        self.dfbLotNumber = lotNumber.strftime("%Y-%m-%d")
        
        print(self.dfbLotNumber)
        print(self.dfbMonth)
        print(self.dfbLetterCode)

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        try:
            if self.dfbYear == "2024":
                self.dfbYear = "2024$"
                vt1Directory = (fr'\\192.168.2.19\{self.dfbYear}')
            else:
                vt1Directory = (fr'\\192.168.2.19\production\{self.dfbYear}')



            for d in os.listdir(vt1Directory):
                if "online checksheet" in d.lower():
                    vt1Directory = os.path.join(vt1Directory, d)
                    for d in os.listdir(vt1Directory):
                        if "outjob" in d.lower():
                            vt1Directory = os.path.join(vt1Directory, d)
                            for d in os.listdir(vt1Directory):
                                if "outjob material monitoring checksheet" in d.lower():
                                    vt1Directory = os.path.join(vt1Directory, d)
                                    os.chdir(vt1Directory)

                                    wb = load_workbook(filename='SNAP.xlsx', data_only=True)

                                    for s in wb.sheetnames:
                                        if self.dfbMonth.lower() in s.lower():
                                            sheet = wb[s]
                                            self.dfbSnapData = pd.DataFrame(sheet.values)
                                            self.dfbSnapData = self.dfbSnapData.iloc[6:]
                                            self.dfbSnapData = self.dfbSnapData.replace(r'\s+', '', regex=True)

                                            #Getting DFB Code
                                            self.dfbCode = self.dfbSnapData.iloc[1, 3]
                                            self.dfbCode = self.dfbCode[8:]
                                            self.dfbCode = self.dfbCode[:-28]

                                            #Filtering SNAP Data, That Contains DFB6600600
                                            self.dfbSnapData = self.dfbSnapData[(self.dfbSnapData[1].isin(["DFB6600600"]))]

                                            break
        except:
            print('No DFB6600600 Snap Not Found')

    def ReadDFB6600600(self):
        try:
            #Converting The First Column/Date To String
            self.dfbSnapData.iloc[:, 0] = self.dfbSnapData.iloc[:, 0].astype(str)

            tempDfbSnapData = self.dfbSnapData[(self.dfbSnapData[0].isin([f"{self.dfbLotNumber} 00:00:00"])) & (self.dfbSnapData[2].isin([self.dfbLetterCode]))]
            
            # tempDfbSnapData = self.dfbSnapData

            # print(self.dfbLotNumber)

            self.dfbLotNumber2 = tempDfbSnapData.iloc[:,3].values[0]

            print(f"Dfb Code {self.dfbCode}")
            print(f"Dfb Lot Number {self.dfbLotNumber2}")
        except:
            print("DFB: There's a problem reading lot number")

    def ReadExcel(self):
        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        while not self.fileFinishedReading:
            try:
                vt1Directory = (fr'\\192.168.2.19\quality control\{str(self.readingYear)}')

                for d in os.listdir(vt1Directory):
                    if "supplier" in d.lower():
                        vt1Directory = os.path.join(vt1Directory, d)
                        for d in os.listdir(vt1Directory):
                            if "inspection standard" in d.lower():
                                vt1Directory = os.path.join(vt1Directory, d)
                                for d in os.listdir(vt1Directory):
                                    if "receiving inspection record" in d.lower():
                                        vt1Directory = os.path.join(vt1Directory, d)
                                        try:
                                            for d in os.listdir(vt1Directory):
                                                if "takaishi" in d.lower():
                                                    directory = os.path.join(vt1Directory, d)

                                                    #Finding A Folder That Contains New Trend
                                                    for d in os.listdir(directory):
                                                        if 'new trend' in d.lower():
                                                            directory = os.path.join(directory, d)
                                                            print(f"Updated vt1Directory: {directory}")
                                                            break

                                                    os.chdir(directory)

                                                    files = glob.glob('*DF06600600*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        self.em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        self.em2PData = pd.DataFrame(self.em2PData)
                                                        self.em2PData = self.em2PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"DFB FINDED IN {self.readingYear} NEW TREND")
                                                        self.fileList.append(self.em2PData)
                                        except:
                                            print("NO DATA FOUND IN GAPTEC")

            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                self.fileFinishedReading = True

        for file in self.fileList:
            file.replace('', np.nan, inplace=True)   

    def GettingData(self):
        for fileNum in range(len(self.fileList)):
            self.totalAverage1 = []
            self.totalAverage2 = []
            self.totalAverage3 = []
            self.totalAverage4 = []

            self.totalMinimum1 = []
            self.totalMinimum2 = []
            self.totalMinimum3 = []
            self.totalMinimum4 = []

            self.totalMaximum1 = []
            self.totalMaximum2 = []
            self.totalMaximum3 = []
            self.totalMaximum4 = []

            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of SUPPLIER
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of Supplier
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
            
            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == self.dfbLotNumber2[:-3]]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 10), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

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

                    self.totalAverage1.append(average1)
                    self.totalAverage2.append(average2)
                    self.totalAverage3.append(average3)
                    self.totalAverage4.append(average4)

                    self.totalMinimum1.append(minimum1)
                    self.totalMinimum2.append(minimum2)
                    self.totalMinimum3.append(minimum3)
                    self.totalMinimum4.append(minimum4)

                    self.totalMaximum1.append(maximum1)
                    self.totalMaximum2.append(maximum2)
                    self.totalMaximum3.append(maximum3)
                    self.totalMaximum4.append(maximum4)

                self.totalAverage1 = statistics.mean(self.totalAverage1)
                self.totalAverage2 = statistics.mean(self.totalAverage2)
                self.totalAverage3 = statistics.mean(self.totalAverage3)
                self.totalAverage4 = statistics.mean(self.totalAverage4)

                self.totalMinimum1 = min(self.totalMinimum1)
                self.totalMinimum2 = min(self.totalMinimum2)
                self.totalMinimum3 = min(self.totalMinimum3)
                self.totalMinimum4 = min(self.totalMinimum4)

                self.totalMaximum1 = max(self.totalMaximum1)
                self.totalMaximum2 = max(self.totalMaximum2)
                self.totalMaximum3 = max(self.totalMaximum3)
                self.totalMaximum4 = max(self.totalMaximum4)

                self.totalAverage1 = f"{self.totalAverage1:.2f}"
                self.totalAverage2 = f"{self.totalAverage2:.2f}"
                self.totalAverage3 = f"{self.totalAverage3:.2f}"
                self.totalAverage4 = f"{self.totalAverage4:.2f}"

                self.totalMinimum1 = f"{self.totalMinimum1:.2f}"
                self.totalMinimum2 = f"{self.totalMinimum2:.2f}"
                self.totalMinimum3 = f"{self.totalMinimum3:.2f}"
                self.totalMinimum4 = f"{self.totalMinimum4:.2f}"

                self.totalMaximum1 = f"{self.totalMaximum1:.2f}"
                self.totalMaximum2 = f"{self.totalMaximum2:.2f}"
                self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                self.totalMaximum4 = f"{self.totalMaximum4:.2f}"

                break

            except:
                self.totalAverage1 = "No Data Found"
                self.totalAverage2 = "No Data Found"
                self.totalAverage3 = "No Data Found"
                self.totalAverage4 = "No Data Found"

                self.totalMinimum1 = "No Data Found"
                self.totalMinimum2 = "No Data Found"
                self.totalMinimum3 = "No Data Found"
                self.totalMinimum4 = "No Data Found"

                self.totalMaximum1 = "No Data Found"
                self.totalMaximum2 = "No Data Found"
                self.totalMaximum3 = "No Data Found"
                self.totalMaximum4 = "No Data Found"

        print(f"Selected Total Average: {self.totalAverage1}")
        print(f"Selected Total Average: {self.totalAverage2}")
        print(f"Selected Total Average: {self.totalAverage3}")
        print(f"Selected Total Average: {self.totalAverage4}")

        print(f"Selected Total Minimum: {self.totalMinimum1}")
        print(f"Selected Total Minimum: {self.totalMinimum2}")
        print(f"Selected Total Minimum: {self.totalMinimum3}")
        print(f"Selected Total Minimum: {self.totalMinimum4}")

        print(f"Selected Total Maximum: {self.totalMaximum1}")
        print(f"Selected Total Maximum: {self.totalMaximum2}")
        print(f"Selected Total Maximum: {self.totalMaximum3}")
        print(f"Selected Total Maximum: {self.totalMaximum4}")

# %%
class Tensile():
    tensileData = ""

    rateOfChangeTotalAverage = []
    rateOfChangeTotalMinimum = []
    rateOfChangeTotalMaximum = []

    startForceTotalAverage = []
    startForceTotalMinimum = []
    startForceTotalMaximum = []

    terminatingForceTotalAverage = []
    terminatingForceTotalMinimum = []
    terminatingForceTotalMaximum = []

    readingYear = ""

    fileList = []
    fileFinishedReading = False

    def __init__(self):
        pass
    def ReadExcel(self):
        self.fileList = []

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        while not self.fileFinishedReading:
            try:
                vt1Directory = (fr'\\192.168.2.19\quality control\{str(self.readingYear)}')

                for d in os.listdir(vt1Directory):
                    if "supplier" in d.lower():
                        vt1Directory = os.path.join(vt1Directory, d)
                        for d in os.listdir(vt1Directory):
                            if "inspection standard" in d.lower():
                                vt1Directory = os.path.join(vt1Directory, d)
                                for d in os.listdir(vt1Directory):
                                    if "receiving inspection record" in d.lower():
                                        vt1Directory = os.path.join(vt1Directory, d)
                                        try:
                                            for d in os.listdir(vt1Directory):
                                                if "tensile" in d.lower():
                                                    directory = os.path.join(vt1Directory, d)

                                                    print(f"Updated vt1Directory: {directory}")

                                                    os.chdir(directory)

                                                    files = glob.glob('*DF06600600*.xlsx')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        self.tensileData = workbook.get_sheet_by_name("Rate_Result_List").to_python(skip_empty_area=True)
                                                        self.tensileData = pd.DataFrame(self.tensileData)
                                                        self.tensileData = self.tensileData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"TENSILE FINDED IN {self.readingYear}")
                                                        self.fileList.append(self.tensileData)
                                        except:
                                            print("NO DATA FOUND IN TENSILE")

            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                self.fileFinishedReading = True

        for file in self.fileList:
            file.replace('', np.nan, inplace=True)

    def GettingData(self, lotNo):
        #Skipping 4 Rows
        self.fileList[0] = self.fileList[0].iloc[4:]

        #Filtering Lot Number Row With Lot Number Input
        tensileLotNoFiltered = self.fileList[0][(self.fileList[0].iloc[:, 3].isin([lotNo]))]

        #Averaging The Rate Of Change Column
        rateOfChangeAverage = round(tensileLotNoFiltered.iloc[:, 6].mean() * 100, 1)
        rateOfChangeMin = round(tensileLotNoFiltered.iloc[:, 6].min() * 100, 1)
        rateOfChangeMax = round(tensileLotNoFiltered.iloc[:, 6].max() * 100, 1)

        self.rateOfChangeTotalAverage = f"{rateOfChangeAverage}%"
        self.rateOfChangeTotalMinimum = f"{rateOfChangeMin}%"
        self.rateOfChangeTotalMaximum = f"{rateOfChangeMax}%"

        #Averaging The Start Force Column
        startForceAverage = tensileLotNoFiltered.iloc[:, 10].mean()
        startForceMin = tensileLotNoFiltered.iloc[:, 10].min()
        startForceMax = tensileLotNoFiltered.iloc[:, 10].max()

        self.startForceTotalAverage = f"{startForceAverage:.1f}"
        self.startForceTotalMinimum = f"{startForceMin:.1f}"
        self.startForceTotalMaximum = f"{startForceMax:.1f}"

        #Averaging The Terminating Column
        terminatingForceAverage = tensileLotNoFiltered.iloc[:, 11].mean()
        terminatingForceMin = tensileLotNoFiltered.iloc[:, 11].min()
        terminationForceMax = tensileLotNoFiltered.iloc[:, 11].max()

        self.terminatingForceTotalAverage = f"{terminatingForceAverage:.1f}"
        self.terminatingForceTotalMinimum = f"{terminatingForceMin:.1f}"
        self.terminatingForceTotalMaximum = f"{terminationForceMax:.1f}"
        
        print(f"RATE OF CHANGE\nAVERAGE: {self.rateOfChangeTotalAverage}\nMINIMUM: {self.rateOfChangeTotalMinimum}\nMAXIMUM: {self.rateOfChangeTotalMaximum}")
        print(f"START FORCE\nAVERAGE: {self.startForceTotalAverage}\nMINIMUM: {self.startForceTotalMinimum}\nMAXIMUM: {self.startForceTotalMaximum}")
        print(f"TERMINATING FORCE\nAVERAGE: {self.terminatingForceTotalAverage}\nMINIMUM: {self.terminatingForceTotalMinimum}\nMAXIMUM: {self.terminatingForceTotalMaximum}")

# %%
# dfb = dFB()
# DateAndTimeManager.GetDateToday()
# dfb.readingYear = int(DateAndTimeManager.yearNow)
# dfb.ReadDfbSnap("20241031-A")
# dfb.ReadDFB6600600()
# dfb.ReadExcel()
# dfb.GettingData()


# %%
# tensile = Tensile()
# DateAndTimeManager.GetDateToday()
# tensile.readingYear = int(DateAndTimeManager.yearNow)
# tensile.ReadExcel()
# tensile.GettingData("T000727")



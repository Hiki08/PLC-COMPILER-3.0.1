#%%
from Imports import *
import DateAndTimeManager

#%%
class cSB():
    csbData = ""
    csbItemCode = ""

    totalAverage1 = []
    
    totalMinimum1 = []

    totalMaximum1 = []
    
    readingYear = ""
    fileFinishedReading = False
    fileList = []

    def __init__(self):
        pass
    def ReadExcel(self, itemCode):
        self.csbItemCode = itemCode

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
                                        if itemCode == "CSB6400802":
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

                                                        files = glob.glob('*CSB6400802*.xlsm')

                                                        for f in files:
                                                            print(f'File Readed {f}')
                                                            workbook = CalamineWorkbook.from_path(f)

                                                            self.csbData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                            self.csbData = pd.DataFrame(self.csbData)
                                                            self.csbData = self.csbData.replace(r'\s+', '', regex=True)
                                                            
                                                            print(f"CSB FINDED IN {self.readingYear} NEW TREND")
                                                            self.fileList.append(self.csbData)
                                            except:
                                                print("NO DATA FOUND IN CRONICS")

            except:
                pass
            
            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                self.fileFinishedReading = True   

    def GettingData(self, lotNumber):
        for fileNum in range(len(self.fileList)):
            self.totalAverage1 = []

            self.totalMinimum1 = []

            self.totalMaximum1 = []

            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of SUPPLIER
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of Hiblow
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]

            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    # Get The Neighboring Data of Lot Number
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 10), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    #CHECKING THE ITEM CODE
                    if self.csbItemCode == "CSB6400802":
                        average1 = inspectionData.iloc[3].mean()

                        minimum1 = inspectionData.iloc[3].min()

                        maximum1 = inspectionData.iloc[3].max()

                        self.totalAverage1.append(average1)

                        self.totalMinimum1.append(minimum1)

                        self.totalMaximum1.append(maximum1)
                
                if self.csbItemCode == "CSB6400802":
                    self.totalAverage1 = statistics.mean(self.totalAverage1)

                    self.totalMinimum1 = min(self.totalMinimum1)

                    self.totalMaximum1 = max(self.totalMaximum1)

                    self.totalAverage1 = f"{self.totalAverage1:.2f}"

                    self.totalMinimum1 = f"{self.totalMinimum1:.2f}"

                    self.totalMaximum1 = f"{self.totalMaximum1:.2f}"

                    break

            except:
                self.totalAverage1 = "No Data Found"

                self.totalMinimum1 = "No Data Found"

                self.totalMaximum1 = "No Data Found"

        print(f"Selected Total Average: {self.totalAverage1}")

        print(f"Selected Total Minimum: {self.totalMinimum1}")

        print(f"Selected Total Maximum: {self.totalMaximum1}")
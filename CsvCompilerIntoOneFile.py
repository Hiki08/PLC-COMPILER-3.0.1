# %%
from Imports import *
import DateAndTimeManager

# %%
class CsvCompiler:
    global gui

    csvFiles = ""
    dataFrame = []

    def __init__(self):
        pass

    def GettingFiles(self):
        self.csvFiles = ""
        self.dataFrame = []

        compiledCsvDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess')
        os.chdir(compiledCsvDirectory)

        self.csvFiles = glob.glob('*CompiledProcess*.csv')

        for files in self.csvFiles:
            df = pd.read_csv(files)
            self.dataFrame.append(df)

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

        self.dataFrame = pd.concat(self.dataFrame, ignore_index=True)

    def WriteCompiledCsv(self):
        fileDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess')
        os.chdir(fileDirectory)
        print(os.getcwd())

        print("Creating New File")
        #Create Excel File
        newValue = pd.concat([self.dataFrame], axis = 0, ignore_index = True)
        wireFrame = newValue
        wireFrame.to_csv(f"FC1DataBase.csv", index = False)

        gui.finishedCompiling()

    def Trial(self):
        coolDown = False
        readCount = 0

        while gui.isAutoRun:
            print("Auto Run Activated")
            DateAndTimeManager.GetTimeNow()
            print(f"Time Now: {DateAndTimeManager.timeNow}")

            hour = gui.time_picker.hours()
            minutes = gui.time_picker.minutes()
            period = gui.time_picker.period()

            timeSet = f"{hour}:{minutes} {period}"
            timeSet = datetime2.strptime(timeSet, "%I:%M %p")
            timeSet = timeSet.strftime("%H:%M")

            print(f"Time Set: {timeSet}")

            if DateAndTimeManager.timeNow == timeSet and not coolDown:
                coolDown = True
                StartProgram()
                time.sleep(70)
                coolDown = False

            #Clearing Cmd Logs When Reaches 10 Lines
            readCount += 1
            if readCount >= 10:
                os.system('cls')
                readCount = 0

            time.sleep(1)
        print("Auto Run Deactivated")

# %%
def Start():
    global csvCompiler

    csvCompiler = CsvCompiler()
    csvCompiler.GettingFiles()
    csvCompiler.WriteCompiledCsv()

# %%
def StartProgram():
    gui.loading()
    
    #Starting Thread
    threading.Thread(target=Start).start()

# %%
class GUI(tk.Tk):
    isAutoRun = False
    
    def __init__(self):
        super().__init__()
        self.title('FC1 Compiler')
        self.iconbitmap('Icons/HiblowLogo.ico')
        self.geometry('600x500+50+50')
        self.resizable(False, False)

        self.on = PhotoImage(file = "Icons/on.png")
        self.off = PhotoImage(file = "Icons/off.png")
        
    def creatingFrames(self):
        #Frames 1
        self.frame1 = tk.Frame(self)
        self.frame1.pack()

        # configure the grid
        self.frame1.columnconfigure(0, weight=1)
        self.frame1.columnconfigure(1, weight=1)
        
        #Frames 2
        self.frame2 = tk.Frame(self)
        self.frame2.forget()

        # configure the grid
        self.frame2.columnconfigure(0, weight=1)
        self.frame2.columnconfigure(1, weight=1)

    def frame1Content(self):
        # place a label on the root window
        message = tk.Label(self.frame1, text="Compiled CSV\nTo One File", font=("Arial", 12, "bold"))
        message.grid(column=0, row=0, columnspan=2, padx=220)

        # button
        self.compileButton = tk.Button(self.frame1, text='COMPILE', font=("Arial", 12), command = StartProgram, width=18, height=1)
        self.compileButton.grid(column=0, row=1, ipadx=5, ipady=5, pady=10, columnspan=2)
        self.compileButton.config(bg="lightgreen", fg="black")

        self.autoRunLabel = tk.Label(self.frame1, text="Auto Run", font=("Arial", 12, "bold"))
        self.autoRunLabel.grid(column=0, row=2)

        self.autoRunButton = tk.Button(self.frame1, image = self.off, bd = 0, font=("Arial", 12), command = self.toggleAutoRun)
        self.autoRunButton.grid(column=1, row=2, ipadx=5, ipady=5, pady=10)

        self.configureButton = tk.Button(self.frame1, text='CONFIGURE', font=("Arial", 8), command = self.configureTime, width=10, height=1)
        self.configureButton.grid(column=0, row=3, ipadx=5, ipady=5, pady=10, columnspan=2)
        self.configureButton.config(bg="lightgreen", fg="black")

    def frame2Content(self):
        # button
        self.backButton = tk.Button(self.frame2, text='BACK', font=("Arial", 8), command = self.backToHome, width=10, height=1)
        self.backButton.grid(column=0, row=0, ipadx=5, ipady=5, pady=10)
        self.backButton.config(bg="lightgreen", fg="black")

        # button
        self.applyButton = tk.Button(self.frame2, text='APPLY', font=("Arial", 8), command = "self.backToHome", width=10, height=1)
        self.applyButton.grid(column=1, row=0, ipadx=5, ipady=5, pady=10)
        self.applyButton.config(bg="lightgreen", fg="black")

        # place a label on the root window
        message = tk.Label(self.frame2, text="Configure Time", font=("Arial", 12, "bold"))
        message.grid(column=0, row=1, columnspan=2, padx=220)

        self.time_picker = AnalogPicker(self.frame2)
        self.time_picker.grid(column = 0, row = 2, columnspan = 2)
        theme = AnalogThemes(self.time_picker)
        theme.setNavyBlue()

    def loading(self):
        self.compileButton.config(text="Loading...")
        self.compileButton.config(state="disabled")

    def finishedCompiling(self):
        self.compileButton.config(text = "Compiled Successfully")
        time.sleep(3)
        self.compileButton.config(text = "COMPILE")
        self.compileButton.config(state="normal")

    def configureTime(self):
        self.frame1.forget()
        self.frame2.pack()

    def backToHome(self):
        self.frame1.pack()
        self.frame2.forget()

    def toggleAutoRun(self):
        if self.isAutoRun == False:
            self.isAutoRun = True
            self.autoRunButton.config(image = self.on)
            threading.Thread(target=csvCompiler.Trial).start()
        else:
            self.isAutoRun = False
            self.autoRunButton.config(image = self.off)

# %%
csvCompiler = CsvCompiler()

gui = GUI()
gui.creatingFrames()
gui.frame1Content()
gui.frame2Content()

gui.mainloop()



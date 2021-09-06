from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showwarning, showinfo
from tkinter.ttk import *
import os
from shutil import copyfile


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Μετονομασία αρχείων με ΑΦΜ")
        self.window.resizable(False, False)
        self.create_widgets()

    def getInputDir(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχεία για μετονομασία")

        if dName == "":
            return

        self.inputDir.set(dName)
        self.ntrOutputDir.configure(state='readonly')
        self.btnOpenOutputDir.configure(state='normal')

    def getOutputDir(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία")

        if dName == "":
            return

        self.outputDir.set(dName)
        self.btnRun.configure(state='normal')

    def renameToAFM(self):
        inputDir = self.inputDir.get()
        outputDir = self.outputDir.get()

        for path, dirs, files in os.walk(inputDir):
            for file in files:
                if '.' in file:
                    dotPos = file.rindex('.')
                    if '[' in file and ']' in file:
                        startPos = file.index('[')
                        endPos = file.index(']')
                        afm = file[startPos + 1:endPos]
                        ext = file[dotPos:]

                        subdir = path.replace(inputDir, "").lstrip("\\")
                        destinationDir = os.path.join(outputDir, subdir)

                        if not os.path.exists(destinationDir):
                            os.makedirs(destinationDir)

                        target_name = afm + ext
                        copyfile(os.path.join(path, file), os.path.join(destinationDir, target_name))

    def run(self):
        self.btnOpenInputDir.configure(state='disabled')
        self.btnOpenOutputDir.configure(state='disabled')
        self.btnRun.configure(state='disabled')

        self.renameToAFM()

        showinfo(title='Ολοκλήρωση Εκτέλεσης',
                 message='Η μετονομασία των αρχείων ολοκληρώθηκε.')

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lInputDir = Label(self.fData, text="Φάκελος εισόδου:")
        self.lInputDir.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.inputDir = StringVar()
        self.inputDir.set('')
        self.ntrInputDir = Entry(self.fData, width=128, textvariable=self.inputDir)
        self.ntrInputDir.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenInputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getInputDir)
        self.btnOpenInputDir.grid(column=2, row=1, padx=10, pady=10)

        self.lOutputDir = Label(self.fData, text="Φάκελος εξόδου:")
        self.lOutputDir.grid(column=0, row=2, padx=10, pady=10, sticky=E)

        self.outputDir = StringVar()
        self.outputDir.set('')
        self.ntrOutputDir = Entry(self.fData, width=128, state='disabled', textvariable=self.outputDir)
        self.ntrOutputDir.grid(column=1, row=2, padx=10, pady=10, sticky=W)

        self.btnOpenOutputDir = Button(self.fData, text="Επιλέξτε φάκελο...", state='disabled',
                                       command=self.getOutputDir)
        self.btnOpenOutputDir.grid(column=2, row=2, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

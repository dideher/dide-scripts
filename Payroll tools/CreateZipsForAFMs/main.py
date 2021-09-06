import os.path
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo
from os import walk
from zipfile import ZipFile


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Δημιουργία zip αρχείων ανά ΑΦΜ")
        self.window.resizable(False, False)
        self.create_widgets()

    def getOutputDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/",
                                        title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία zip")

        if dName == "":
            return

        self.outputDirName.set(dName)
        self.btnRun.configure(state='normal')

    def getInputDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/",
                                        title="Επιλέξτε τον φάκελο με τα αρχεία εισόδου")

        if dName == "":
            return

        self.inputDirName.set(dName)
        self.ntrOutputDirName.configure(state='readonly')
        self.btnOpenOutputDir.configure(state='normal')

    def verify_afm(self, value):
        afm = value

        if len(afm) != 9:
            return False

        if not afm.isdigit():
            return False

        chcknumbers = [0, 2, 4, 8, 16, 32, 64, 128, 256]
        lchcknumbers = len(chcknumbers) - 1
        sum = 0

        for i in range(9):
            sum += (int(afm[i]) * chcknumbers[lchcknumbers - i])

        ch_digit = int(afm[8])

        ypoloipo = sum % 11

        if ypoloipo == 10:
            ypoloipo = 0

        if ypoloipo == ch_digit:
            return True
        else:
            return False

    def getAFM(self, text):
        split_text = text.split('.')

        return split_text[0][-9:]

    def createAFMlist(self):
        self.afms = set()
        self.errors = ''
        self.err_counter = 0

        for path, dirs, files in walk(self.inputDirName.get()):
            for file in files:
                afm = self.getAFM(file)

                if self.verify_afm(afm):
                    self.afms.add(afm)
                else:
                    self.err_counter += 1
                    self.errors += f'{self.err_counter}: {path}/{file}\n'

    def createZips(self):
        inputDir = self.inputDirName.get()

        for afm in self.afms:

            zipObj = ZipFile(os.path.join(self.outputDirName.get(), f'{afm}.zip'), 'w')

            for path, dirs, files in walk(inputDir):
                for file in files:
                    if afm in file:
                        file_path = os.path.join(path, file)
                        zip_path = file_path.replace(inputDir, '')
                        zipObj.write(file_path, zip_path)

            zipObj.close()

    def run(self):
        self.btnRun.configure(state='disabled')

        self.createAFMlist()
        self.createZips()

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="H δημιουργία των zip αρχείων ολοκληρώθηκε.")

        if self.errors != '':
            print(20 * '-', 'Λάθη', 20 * '-')
            print(self.errors)

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lInputDirName = Label(self.fData, text="Φάκελος με αρχεία εισόδου:")
        self.lInputDirName.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.inputDirName = StringVar()
        self.ntrInputDirName = Entry(self.fData, width=128, state='readonly', textvariable=self.inputDirName)
        self.ntrInputDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenInputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getInputDirName)
        self.btnOpenInputDir.grid(column=2, row=0, padx=10, pady=10)

        self.lOutputDirName = Label(self.fData, text="Φάκελος για αποθήκευση των αρχείων:")
        self.lOutputDirName.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.outputDirName = StringVar()
        self.ntrOutputDirName = Entry(self.fData, width=128, state='disabled', textvariable=self.outputDirName)
        self.ntrOutputDirName.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenOutputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getOutputDirName,
                                       state='disabled')
        self.btnOpenOutputDir.grid(column=2, row=1, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

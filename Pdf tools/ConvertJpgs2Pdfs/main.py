from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showinfo
from tkinter.ttk import *
from os import walk
from img2pdf import convert


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Μετατροπή αρχείων jpg σε pdf")
        self.window.resizable(False, False)
        self.create_widgets()


    def getFilesDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχεία προς επεξεργασία")

        if dName == "":
            return

        self.filesDirName.set(dName)
        self.ntrOutputDirName.configure(state='normal')
        self.btnOpenOutputDir.configure(state='normal')


    def getOutputDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία")

        if dName == "":
            return

        self.outputDirName.set(dName)
        self.btnRun.configure(state='normal')


    def getAFMfromFile(self, filename):
        return filename.replace(".jpg", "").split('_')[0]


    def run(self):
        self.btnRun.configure(state='disabled')

        inputDirectory = self.filesDirName.get()
        outputDirectory = self.outputDirName.get()

        afms = dict()
        for path, dirs, files in walk(inputDirectory):
            for file in files:
                if file[-4:] == ".jpg":
                    afm = self.getAFMfromFile(file)

                    if afm not in afms:
                        afms[afm] = list()

                    afms[afm].append(path + '/' + file)

        files_count = len(afms)

        self.pbProgress['maximum'] = files_count
        self.pbProgress['value'] = 0
        self.pbProgress.update()

        count = 0
        for afm in afms:
            self.lProgress.configure(text=f'Εκτέλεση σε εξέλιξη... ({count + 1}/{files_count})')
            self.pbProgress['value'] = count + 1
            self.pbProgress.update()

            pdfdata = convert(afms[afm])

            pdfFile = open(f'{outputDirectory}/{afm}.pdf', 'wb')
            pdfFile.write(pdfdata)
            pdfFile.close()

            count += 1

        showinfo(title="Ολοκλήρωση επεξεργασίας", message=f'Η μετατοπή των αρχείων jpg σε pdf ολοκληρώθηκε.')


    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lData = Label(self.fData, text="Φάκελος με αρχεία jpg:")
        self.lData.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.filesDirName = StringVar()
        self.ntrFilesDirName = Entry(self.fData, width=100, state='readonly', textvariable=self.filesDirName)
        self.ntrFilesDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenData = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getFilesDirName)
        self.btnOpenData.grid(column=2, row=0, padx=10, pady=10)

        self.lOutputDirName = Label(self.fData, text="Φάκελος για αποθήκευση των pdf:")
        self.lOutputDirName.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.outputDirName = StringVar()
        self.ntrOutputDirName = Entry(self.fData, width=100, state='disabled', textvariable=self.outputDirName)
        self.ntrOutputDirName.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenOutputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getOutputDirName,
                                       state='disabled')
        self.btnOpenOutputDir.grid(column=2, row=1, padx=10, pady=10)

        self.lProgress = Label(self.fData, text="Αναμονή για εκτέλεση ...")
        self.lProgress.grid(column=0, row=2, columnspan=3, padx=10, pady=10)

        self.pbProgress = Progressbar(self.fData, orient='horizontal', length=400, mode='determinate')
        self.pbProgress.grid(column=0, row=3, columnspan=3, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=0, row=10, columnspan=3, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

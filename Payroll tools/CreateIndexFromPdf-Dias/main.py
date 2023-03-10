from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showwarning, showinfo
import PyPDF2
import re
import csv


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Δημιουργία Ευρετηρίου Βεβαιώσεων Αποδοχών")
        self.window.resizable(False, False)
        self.create_widgets()


    def createPdf(self):
        inFile = self.dataFilename.get()
        pdfFileObj = open(inFile, 'rb')  # 'rb' for read binary mode
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

        dataList = list()
        dataList.append(["afm", "amka", "page", "length of afm", "length of amka"])

        self.pbProgress['maximum'] = pdfReader.numPages
        self.pbProgress['value'] = 0
        self.pbProgress.update()

        for i in range(pdfReader.numPages):
            self.lProgress.configure(text=f'Εκτέλεση σε εξέλιξη... ({i + 1}/{pdfReader.numPages})')
            self.pbProgress['value'] = i + 1
            self.pbProgress.update()

            pageObj = pdfReader.getPage(i)
            pageText = re.sub(r'\s\s+', ' ', pageObj.extractText().replace("\n", " ").strip()).split(" ")

            afm = ''
            amka = ''
            afmFound = False
            amkaFound = False

            for item in pageText:
                if (len(item) == 9 and item.isdigit()):
                    afm = item
                    afmFound = True
                if (len(item) == 11 and item.isdigit()):
                    amka = item
                    amkaFound = True
                if (afmFound and amkaFound):
                    break

            dataList.append([afm, amka, i + 1, len(afm), len(amka)])
            print(f'{i + 1}: {afm}, {amka}')

        with open("emp.csv", 'w', newline='') as myfile:
            wr = csv.writer(myfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            for entry in dataList:
                wr.writerow(entry)

    def getDataFilename(self):
        fName = filedialog.askopenfilename(initialdir="./data/",
                                           title="Επιλέξτε το αρχείο pdf",
                                           filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

        if fName == "":
            return

        self.dataFilename.set(fName)
        self.btnRun.configure(state='normal')

    def run(self):
        self.btnRun.configure(state='disabled')
        self.createPdf()
        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η δημιουργία του Ευρετηρίου Βεβαιώσεων Αποδοχών ολοκληρώθηκε.")

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lData = Label(self.fData, text="Αρχείο:")
        self.lData.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.dataFilename = StringVar()
        self.ntrDataFilename = Entry(self.fData, width=128, state='readonly', textvariable=self.dataFilename)
        self.ntrDataFilename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenData = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getDataFilename)
        self.btnOpenData.grid(column=2, row=0, padx=10, pady=10)

        self.lProgress = Label(self.fData, text="Αναμονή για εκτέλεση ...")
        self.lProgress.grid(column=0, row=1, columnspan=3, padx=10, pady=10)

        self.pbProgress = Progressbar(self.fData, orient='horizontal', length=400, mode='determinate')
        self.pbProgress.grid(column=0, row=2, columnspan=3, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=0, row=10, columnspan=3, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

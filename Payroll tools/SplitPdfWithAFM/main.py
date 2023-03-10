from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo
import fitz
import os


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title("Διαχωρισμός αρχείου pdf ανά ΑΦΜ")
        self.window.resizable(False, False)
        self.create_widgets()

    def getOutputDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία")

        if dName == "":
            return

        self.outputDirName.set(dName)
        self.btnRun.configure(state='normal')

    def getDataFilename(self):
        fName = filedialog.askopenfilename(initialdir="./data/",
                                           title="Επιλέξτε το αρχείο pdf",
                                           filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

        if fName == "":
            return

        self.dataFilename.set(fName)
        self.ntrOutputDirName.configure(state='readonly')
        self.btnOpenOutputDir.configure(state='normal')

    def removeSpaces(self, text_list):
        while '' in text_list:
            text_list.remove('')

        return text_list

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

    def findAFM(self, l):
        afmList = list()

        for item in l:
            if self.verify_afm(item) and item != '099709779':
                afmList.append(item)

        if afmList:
            return afmList[-1]
        else:
            return self.last_afm

    def run(self):
        self.btnRun.configure(state='disabled')

        filename = self.dataFilename.get()
        outputDir = self.outputDirName.get()

        data = dict()

        pdf = fitz.open(filename)
        self.last_afm = '000000000'
        for page in range(pdf.page_count):
            pageObj = pdf[page]
            pageText = pageObj.get_text("text")
            textList = self.removeSpaces(pageText.replace('\n', ' ').split(' '))

            afm = self.findAFM(textList)
            self.last_afm = afm
            if afm not in data:
                data[afm] = list()
            data[afm].append(page)

        for afm in data:
            pdf_out = fitz.open()

            for page in data[afm]:
                pdf_out.insert_pdf(pdf, from_page=page, to_page=page)

            output_pdf = os.path.join(outputDir, afm + ".pdf")
            pdf_out.save(output_pdf)
            pdf_out.close()

        pdf.close()

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Ο διαχωρισμός ολοκληρώθηκε.")

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lData = Label(self.fData, text="Αρχείο:")
        self.lData.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.dataFilename = StringVar()
        self.ntrDataFilename = Entry(self.fData, width=128, state='readonly', textvariable=self.dataFilename)
        self.ntrDataFilename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenData = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getDataFilename)
        self.btnOpenData.grid(column=2, row=0, padx=10, pady=10)

        self.lOutputDirName = Label(self.fData, text="Φάκελος για αποθήκευση των αρχείων:")
        self.lOutputDirName.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.outputDirName = StringVar()
        self.ntrOutputDirName = Entry(self.fData, width=128, state='disabled', textvariable=self.outputDirName)
        self.ntrOutputDirName.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenOutputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getOutputDirName,
                                       state='disabled')
        self.btnOpenOutputDir.grid(column=2, row=1, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση διαχωρισμού", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

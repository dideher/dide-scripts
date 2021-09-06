from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo
from PyPDF2 import PdfFileReader, PdfFileWriter
import slate3k as slate
import os


class GUI():
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

        return afmList[-1]

    def run(self):
        self.btnRun.configure(state='disabled')

        filename = self.dataFilename.get()
        outputDir = self.outputDirName.get()

        data = dict()

        showinfo(title="Έναρξη εκτέλεσης",
                 message="Η ανάγνωση του κειμένου απαιτεί αρκετό χρόνο.\nMην τερματίσετε την εφαρμογή μέχρι να εμφανιστεί το μήνυμα της ολοκλήρωσης.")

        extracted_text = slate.PDF(open(filename, 'rb'))
        pdf = PdfFileReader(filename)
        for page in range(pdf.getNumPages()):
            pageObj = pdf.getPage(page)
            pageText = self.removeSpaces(extracted_text[page].replace("\n", " ").strip().split(" "))

            afm = self.findAFM(pageText)
            if afm not in data:
                data[afm] = list()
            data[afm].append(pageObj)

        for afm in data:
            pdf_writer = PdfFileWriter()

            for pageObj in data[afm]:
                pdf_writer.addPage(pageObj)

            output = os.path.join(outputDir, afm + ".pdf")
            with open(output, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)

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

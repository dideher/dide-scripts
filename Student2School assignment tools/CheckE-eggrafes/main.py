from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showwarning, showinfo
from tkinter.ttk import *
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import csv


class GUI():
    def __init__(self):
        self.schoolsOfInterestInCsv = [
            "10ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "11ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "13ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "1ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "2ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟ",
            "3ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "4ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "5ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "6ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "7ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "8ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΓΑΖΙΟΥ ΗΡΑΚΛΕΙΟ ΚΡΗΤΗΣ - ΔΟΜΗΝΙΚΟΣ ΘΕΟΤΟΚΟΠΟΥΛΟΣ",
            "ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΝΕΑΣ ΑΛΙΚΑΡΝΑΣΣΟΥ"
        ]

        self.schoolsOfInterestInXlsx = [
            "10ο ΓΕΛ",
            "11ο ΓΕΛ",
            "13ο ΓΕΛ",
            "1ο ΓΕΛ",
            "2ο ΓΕΛ",
            "3ο ΓΕΛ",
            "4ο ΓΕΛ",
            "5ο ΓΕΛ",
            "6ο ΓΕΛ",
            "7ο ΓΕΛ",
            "8ο ΓΕΛ",
            "ΓΕΛ ΓΑΖΙΟΥ",
            "ΓΕΛ ΝΕΑΣ ΑΛΙΚΑΡΝΑΣΣΟΥ"
        ]

        self.shortToFullLectic = {
            "10ο ΓΕΛ": "10ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "11ο ΓΕΛ": "11ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "13ο ΓΕΛ": "13ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "1ο ΓΕΛ": "1ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "2ο ΓΕΛ": "2ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟ",
            "3ο ΓΕΛ": "3ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "4ο ΓΕΛ": "4ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "5ο ΓΕΛ": "5ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "6ο ΓΕΛ": "6ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "7ο ΓΕΛ": "7ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "8ο ΓΕΛ": "8ο ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΗΡΑΚΛΕΙΟΥ",
            "ΓΕΛ ΓΑΖΙΟΥ": "ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΓΑΖΙΟΥ ΗΡΑΚΛΕΙΟ ΚΡΗΤΗΣ - ΔΟΜΗΝΙΚΟΣ ΘΕΟΤΟΚΟΠΟΥΛΟΣ",
            "ΓΕΛ ΝΕΑΣ ΑΛΙΚΑΡΝΑΣΣΟΥ": "ΗΜΕΡΗΣΙΟ ΓΕΝΙΚΟ ΛΥΚΕΙΟ ΝΕΑΣ ΑΛΙΚΑΡΝΑΣΣΟΥ"
        }

        self.gymsToIgnore = [
            'ΗΜΕΡΗΣΙΟ ΓΥΜΝΑΣΙΟ ΤΥΛΙΣΟΥ "ΙΩΑΝΝΗΣ ΠΕΡΔΙΚΑΡΗΣ"',
            'ΗΜΕΡΗΣΙΟ ΓΥΜΝΑΣΙΟ ΝΕΑΣ ΑΛΙΚΑΡΝΑΣΣΟΥ',
            'ΗΜΕΡΗΣΙΟ ΓΥΜΝΑΣΙΟ ΠΡΟΦΗΤΗ ΗΛΙΑ ΗΡΑΚΛΕΙΟΥ'
        ]
        self.window = Tk()

        self.window.title("Επαλήθευση καταχωρίσεων Λυκείων πόλης στο e-eggrafes")
        self.window.resizable(False, False)
        self.create_widgets()

    def getCsvFilename(self):
        fName = filedialog.askopenfilename(initialdir=".", title="Επιλέξτε το αρχείο csv",
                                           filetypes=(("csv files", "*.csv"), ("all files", "*.*")))

        if fName == "":
            return

        self.btnOpenCsvFile.configure(state='disabled')
        self.csvFilename.set(fName)

        self.csvData = self.parseCsvData(fName)
        self.csvData = self.cleanQuotes(self.csvData)
        self.ntrXlsxFilename.configure(state='readonly')
        self.btnOpenXlsxFile.configure(state='normal')

    def getXlsxFilename(self):
        fName = filedialog.askopenfilename(initialdir=".", title="Επιλέξτε το αρχείο xlsx",
                                           filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

        if fName == "":
            return

        self.btnOpenXlsxFile.configure(state='disabled')
        self.xlsxFilename.set(fName)

        self.xlsxData = self.parseXlsxData(fName)
        self.btnRun.configure(state='normal')

    def parseXlsxData(self, fileName):
        workbook = load_workbook(filename=fileName)
        sheet = workbook.active

        data = list()

        for row in sheet.iter_rows():
            entry = list()
            for cell in row:
                if cell.value is None:
                    entry.append("")
                else:
                    text1 = (
                        str(cell.value).upper().replace(".", ". ").replace(" .", ". ").replace("Ά", "Α").replace(
                            "Έ", "Ε")
                            .replace("Ή", "Η").replace("Ί", "Ι").replace("Ϊ́", "Ϊ").replace("Ύ", "Υ").replace(
                            "Ϋ́",
                            "Ϋ")
                            .replace("Ό", "Ο").replace("Ώ", "Ω").strip())

                    text2 = re.sub(r'([ ]+)', r' ', text1)
                    entry.append(re.sub(r'([0-9]+)Ο', r'\1ο', text2))

            data.append(entry)

        return data[1:]

    def parseCsvData(self, inputFile):
        data = list()

        with open(inputFile, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)

            for row in reader:
                data.append(row)

        return data[1:]

    def cleanQuotes(self, data):
        cleanData = list()

        for entry in data:
            cleanEntry = list()

            for i, item in enumerate(entry):
                cleanItem = item.replace("'", "")

                if i == 0:
                    am_aa = cleanItem.split("/")
                    cleanEntry.append(am_aa[0].strip())
                    cleanEntry.append(am_aa[1].strip())
                else:
                    cleanEntry.append(cleanItem)

            cleanData.append(cleanEntry)

        return cleanData

    def filterList(self, data, index, itemsOfInterest):
        filterData = list()

        for entry in data:
            if entry[index] in itemsOfInterest:
                filterData.append(entry)

        return filterData

    def exludeList(self, data, index, itemsOfInterest):
        filterData = list()

        for entry in data:
            if entry[index] in itemsOfInterest:
                continue
            filterData.append(entry)

        return filterData

    def saveFile(self, data, outputFile):
        wb = Workbook()
        ws = wb.active

        for entry in data:
            ws.append(entry)

        column_widths = []
        for row in ws.iter_rows():
            for i, cell in enumerate(row):
                try:
                    column_widths[i] = max(column_widths[i], len(str(cell.value)))
                except IndexError:
                    column_widths.append(len(str(cell.value)))

        for i, column_width in enumerate(column_widths):
            ws.column_dimensions[get_column_letter(i + 1)].width = column_width * 1.23

        notSaved = True

        while notSaved:
            try:
                wb.save(outputFile)
            except:
                showwarning(title="Αρχείο σε χρήση...",
                            message="Παρακαλώ κλείστε το αρχείο '{}' ώστε να ολοκληρωθεί η αποθήκευση.".format(
                                outputFile))
            else:
                notSaved = False

    def check(self):
        notFoundErrors = ''

        for csvEntry in self.csvFiltered:
            found = False
            for xlsxEntry in self.xlsxFiltered:
                if csvEntry[1] == xlsxEntry[1]:
                    found = True
                    if csvEntry[6] != self.shortToFullLectic[xlsxEntry[6]]:
                        print(f'Ασυμφωνία: E-eggrafes {csvEntry} <--> Κατανομή {xlsxEntry}')
                    break

            if not found:
                notFoundErrors += f'Δεν βρέθηκε: {csvEntry}\n'

        if notFoundErrors != '':
            print(notFoundErrors)

    def run(self):
        self.btnRun.configure(state='disabled')

        self.csvFiltered = self.exludeList(self.csvData, 5, self.gymsToIgnore)
        self.csvFiltered = self.filterList(self.csvFiltered, 6, self.schoolsOfInterestInCsv)
        self.xlsxFiltered = self.filterList(self.xlsxData, 6, self.schoolsOfInterestInXlsx)
        self.check()

        showinfo(title='Ολοκλήρωση Εκτέλεσης',
                 message=f'Ο έλεγχος ολοκληρώθηκε.')

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lCsvFile = Label(self.fData, text="Αρχείο αναφοράς e-eggrafes (csv):")
        self.lCsvFile.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.csvFilename = StringVar()
        self.csvFilename.set('')
        self.ntrCsvFilename = Entry(self.fData, width=128, state='readonly', textvariable=self.csvFilename)
        self.ntrCsvFilename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenCsvFile = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getCsvFilename)
        self.btnOpenCsvFile.grid(column=2, row=0, padx=10, pady=10)

        self.lXlsxFile = Label(self.fData, text="Αρχείο κατανομής (xlsx):")
        self.lXlsxFile.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.xlsxFilename = StringVar()
        self.xlsxFilename.set('')
        self.ntrXlsxFilename = Entry(self.fData, width=128, state='disabled', textvariable=self.xlsxFilename)
        self.ntrXlsxFilename.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenXlsxFile = Button(self.fData, text="Επιλέξτε αρχείο...", state='disabled',
                                      command=self.getXlsxFilename)
        self.btnOpenXlsxFile.grid(column=2, row=1, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

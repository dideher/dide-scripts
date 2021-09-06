from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showwarning, showinfo
from tkinter.ttk import *
from openpyxl import *
from openpyxl.utils import get_column_letter
from parseXlsxData import *


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Διαχωρισμός μαθητών σε πόλης και περιφέρειας")
        self.window.resizable(False, False)
        self.create_widgets()


    def getStudentsFilename(self):
        fName = filedialog.askopenfilename(initialdir=".", title="Επιλέξτε το αρχείο των μαθητών",
                                           filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

        if fName == "":
            return

        self.btnOpenStudentsFile.configure(state='disabled')
        self.studentsFilename.set(fName)
        self.students = parseStudents(fName)
        self.cbSchoolCol['values'] = self.students[0]
        self.cbSchoolCol.configure(state='readonly')
        self.urban = list()
        self.rural = list()


    def getSchoolsFilename(self):
        fName = filedialog.askopenfilename(initialdir=".", title="Επιλέξτε το αρχείο των σχολείων",
                                           filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

        if fName == "":
            return

        self.btnOpenSchoolsFile.configure(state='disabled')
        self.btnRun.configure(state='normal')
        self.schoolsFilename.set(fName)
        self.schools = parseSchools(fName)


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
                showwarning(title="Αρχείο σε χρήση...", message="Παρακαλώ κλείστε το αρχείο '{}' ώστε να ολοκληρωθεί η αποθήκευση.".format(outputFile))
            else:
                notSaved = False


    def run(self):
        self.cbSchoolCol.configure(state='disabled')
        self.btnRun.configure(state='disabled')

        sc = self.cbSchoolCol.current()

        self.urban.append(self.students[0])
        self.rural.append(self.students[0] + ['ΣΧΟΛΕΙΟ ΚΑΤΑΝΟΜΗΣ', ])

        for student in self.students[1:]:
            if student[sc] in self.schools:
                self.rural.append(student + [self.schools[student[sc]], ])
            else:
                self.urban.append(student)

        self.saveFile(self.urban, "urban.xlsx")
        self.saveFile(self.rural, "rural.xlsx")

        showinfo(title="Αρχεία εξόδου",
                    message=f'Τα αποτελέσματα έχουν αποθηκευτεί στον φάκελο εκτέλεσης του προγράμματος.')

    def cbSchoolColSelect(self, eventObject):
        if self.cbSchoolCol.current() != -1:
            self.ntrSchoolsFilename.configure(state='readonly')
            self.btnOpenSchoolsFile.configure(state='normal')

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lStudentsFile = Label(self.fData, text="Αρχείο μαθητών:")
        self.lStudentsFile.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.studentsFilename = StringVar()
        self.studentsFilename.set('')
        self.ntrStudentsFilename = Entry(self.fData, width=128, state='readonly', textvariable=self.studentsFilename)
        self.ntrStudentsFilename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenStudentsFile = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getStudentsFilename)
        self.btnOpenStudentsFile.grid(column=2, row=0, padx=10, pady=10)

        self.lSchoolCol = Label(self.fData, text="Στήλη Σχολείου Προέλευσης:")
        self.lSchoolCol.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.schoolCol = StringVar()
        self.cbSchoolCol = Combobox(self.fData, width=125, textvariable=self.schoolCol, state='disabled')
        self.cbSchoolCol.grid(column=1, row=1, padx=10, pady=10, sticky=W)
        self.cbSchoolCol.bind("<<ComboboxSelected>>", self.cbSchoolColSelect)

        self.lSchoolsFile = Label(self.fData, text="Αρχείο σχολείων\nπεριφέρειας:")
        self.lSchoolsFile.grid(column=0, row=2, padx=10, pady=10, sticky=E)

        self.schoolsFilename = StringVar()
        self.schoolsFilename.set('')
        self.ntrSchoolsFilename = Entry(self.fData, width=128, state='disabled', textvariable=self.schoolsFilename)
        self.ntrSchoolsFilename.grid(column=1, row=2, padx=10, pady=10, sticky=W)

        self.btnOpenSchoolsFile = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getSchoolsFilename, state='disabled')
        self.btnOpenSchoolsFile.grid(column=2, row=2, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση διαχωρισμού", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

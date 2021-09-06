from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showwarning, showinfo
from openpyxl.utils import get_column_letter
from openpyxl import *
import os
import geopy.distance


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Υπολογισμός αποστάσεων")
        self.window.resizable(False, False)

        self.data1 = list()
        self.data2 = list()
        self.data3 = list()

        self.create_widgets()

    def setColsWidth(self, ws):
        column_widths = []
        for row in ws.iter_rows():
            for i, cell in enumerate(row):
                try:
                    column_widths[i] = max(column_widths[i], len(str(cell.value)))
                except IndexError:
                    column_widths.append(len(str(cell.value)))

        for i, column_width in enumerate(column_widths):
            ws.column_dimensions[get_column_letter(i + 1)].width = column_width * 1.23

    def saveFile(self):
        wb = Workbook()
        ws = wb.active

        for row in self.data3:
            ws.append(row)

        self.setColsWidth(ws)

        outputFile = "output.xlsx"

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

    def calc(self):
        for item2 in self.data2[1:]:
            coords_2 = (item2[3], item2[4])
            schools = list()

            for item1 in self.data1[1:]:
                coords_1 = (item1[1], item1[2])

                schools.append([item1[0], str(geopy.distance.distance(coords_1, coords_2).km)])
            schools.sort(key=self.schDistance)

            entry = item2[:3]
            for item in schools:
                entry += [f'{item[0]} ({item[1]})']

            self.data3.append(entry)

        print(self.data3)
        self.saveFile()
        os.startfile("output.xlsx")

    def schDistance(self, l):
        return float(l[1])

    def parseXlsxData1(self):
        workbook = load_workbook(filename=self.dataFilename1.get())
        sheet = workbook.active

        for row in sheet.iter_rows():
            entry = list()
            for cell in row:
                if cell.value is None:
                    text = ""
                else:
                    text = str(cell.value).strip()

                entry.append(text)

            self.data1.append(entry)

    def parseXlsxData2(self):
        workbook = load_workbook(filename=self.dataFilename2.get())
        sheet = workbook.active

        for r, row in enumerate(sheet.iter_rows()):
            entry = list()
            for i, cell in enumerate(row):
                if i not in [1, 2, 3, 11]:
                    continue

                if cell.value is None:
                    text = ""
                else:
                    text = str(cell.value).strip()

                if i == 11:
                    if r == 0:
                        entry.append('Longitude')
                        entry.append('Latitude')
                    else:
                        lat, lng = text.split(',')
                        entry.append(lng)
                        entry.append(lat)
                else:
                    entry.append(text)

            self.data2.append(entry)

    def getDataFilename1(self):
        fName = filedialog.askopenfilename(initialdir="./data/",
                                           title="Επιλέξτε το αρχείο των σχολείων",
                                           filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

        if fName == "":
            return

        self.dataFilename1.set(fName)
        self.ntrDataFilename1.configure(state='disabled')
        self.btnOpenData1.configure(state='disabled')

        self.parseXlsxData1()

        self.ntrDataFilename2.configure(state='readonly')
        self.btnOpenData2.configure(state='normal')

    def getDataFilename2(self):
        fName = filedialog.askopenfilename(initialdir="./data/",
                                           title="Επιλέξτε το αρχείο των μαθητών",
                                           filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

        if fName == "":
            return

        self.dataFilename2.set(fName)
        self.ntrDataFilename2.configure(state='disabled')
        self.btnOpenData2.configure(state='disabled')

        self.parseXlsxData2()

        self.btnRun.configure(state='normal')

    def run(self):
        self.btnRun.configure(state='disabled')
        self.calc()
        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Ο υπολογισμός ολοκληρώθηκε.")

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lData1 = Label(self.fData, text="Αρχείο σχολείων:")
        self.lData1.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.dataFilename1 = StringVar()
        self.ntrDataFilename1 = Entry(self.fData, width=128, state='readonly', textvariable=self.dataFilename1)
        self.ntrDataFilename1.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenData1 = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getDataFilename1)
        self.btnOpenData1.grid(column=2, row=0, padx=10, pady=10)

        self.lData2 = Label(self.fData, text="Αρχείο μαθητών:")
        self.lData2.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.dataFilename2 = StringVar()
        self.ntrDataFilename2 = Entry(self.fData, width=128, state='disabled', textvariable=self.dataFilename2)
        self.ntrDataFilename2.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenData2 = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getDataFilename2,
                                   state='disabled')
        self.btnOpenData2.grid(column=2, row=1, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

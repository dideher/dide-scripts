from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo, showwarning
from PyPDF2 import PdfFileReader, PdfFileWriter
import os
import re


def atoi(text):
    return int(text) if text.isdigit() else text


def natural_keys(text):
    return [ atoi(c) for c in re.split(r'(\d+)', text) ]


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Συνένωση αρχείων pdf")
        self.window.resizable(False, False)
        self.create_widgets()


    def getInputDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που περιέχει τα αρχεία")

        if dName == "":
            return

        if dName != self.inputDirName.get():
            self.notMerged = True

        self.inputDirName.set(dName)
        self.btnRun.configure(state='normal')


    def run(self):
        if not self.notMerged:
            showwarning(title="Προσοχή ...", message="Έχετε εκτελέσει ήδη τη συνένωση των αρχείων αυτού του φακέλου.")
            return

        inputDir = self.inputDirName.get()
        filesCounter = 0

        pdf_writer = PdfFileWriter()

        if self.numbersChecked.get() == 1:
            files = list()

            for entry in os.scandir(inputDir):
                if entry.is_file() and entry.name[-4:] == ".pdf":
                    files.append(entry.name)

            files.sort(key=natural_keys)

            for file in files:
                filesCounter += 1
                pdf = PdfFileReader(os.path.join(inputDir, file))

                for page in range(pdf.getNumPages()):
                    pdf_writer.addPage(pdf.getPage(page))
        else:
            for entry in os.scandir(inputDir):
                if entry.is_file() and entry.name[-4:] == ".pdf":
                    filesCounter += 1

                    pdf = PdfFileReader(os.path.join(inputDir, entry.name))

                    for page in range(pdf.getNumPages()):
                        pdf_writer.addPage(pdf.getPage(page))

        if filesCounter == 0:
            showwarning(title="Αποτυχία εκτέλεσης", message="Δεν υπήρχαν αρχεία για συνένωση.")
        else:
            output = os.path.join(inputDir, "output.pdf")
            with open(output, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)

            self.notMerged = False
            showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η συνένωση ολοκληρώθηκε.")


    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lInputDirName = Label(self.fData, text="Φάκελος αρχείων:")
        self.lInputDirName.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.inputDirName = StringVar()
        self.ntrInputDirName = Entry(self.fData, width=128, state='disabled', textvariable=self.inputDirName)
        self.ntrInputDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenInputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getInputDirName)
        self.btnOpenInputDir.grid(column=2, row=0, padx=10, pady=10)

        self.numbersChecked = IntVar()
        self.ckbNumbersChecked = Checkbutton(self.fData, text="Αριθμητική ταξινόμηση αρχείων", variable=self.numbersChecked)
        self.ckbNumbersChecked.grid(column=1, row=1, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση συνένωσης", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

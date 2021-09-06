from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo
import os
from win32com import client


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Μετατροπή αρχείων docx σε pdf")
        self.window.resizable(False, False)
        self.create_widgets()

    def getInputDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που περιέχει τα αρχεία docx")

        if dName == "":
            return

        self.inputDirName.set(dName)
        self.ntrOutputDirName.configure(state='readonly')
        self.btnOpenOutputDir.configure(state='normal')

    def getOutputDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/",
                                        title="Επιλέξτε τον φάκελο που θα αποθηκεύσετε τα αρχεία pdf")

        if dName == "":
            return

        self.outputDirName.set(dName)
        self.btnRun.configure(state='normal')

    def run(self):
        self.btnRun.configure(state='disabled')
        self.word = client.Dispatch('Word.Application')

        inputDir = self.inputDirName.get()
        outputDir = self.outputDirName.get()

        try:
            word = client.Dispatch("Word.Application")

            for file in os.listdir(inputDir):
                in_file = os.path.join(inputDir, file)

                if file.endswith(".docx") and os.path.isfile(in_file):
                    out_file = os.path.join(outputDir, file.replace(".docx", ".pdf"))
                    doc = word.Documents.Open(in_file)

                    if os.path.exists(out_file):
                        os.remove(out_file)

                    doc.SaveAs(out_file, FileFormat=17)
                    doc.Close()

            word.Quit()
        except Exception:
            print("Exception !!!")

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η μετατροπή ολοκληρώθηκε.")

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lInputDirName = Label(self.fData, text="Φάκελος αρχείων docx:")
        self.lInputDirName.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.inputDirName = StringVar()
        self.ntrInputDirName = Entry(self.fData, width=128, state='readonly', textvariable=self.inputDirName)
        self.ntrInputDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenInputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getInputDirName)
        self.btnOpenInputDir.grid(column=2, row=0, padx=10, pady=10)

        self.lOutputDirName = Label(self.fData, text="Φάκελος αρχείων pdf:")
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

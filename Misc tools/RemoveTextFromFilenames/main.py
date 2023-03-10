from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showwarning, showinfo
from tkinter.ttk import *
import os
from shutil import copyfile


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Αφαίρεση κειμένου από το όνομα των αρχείων")
        self.window.resizable(False, False)
        self.create_widgets()

    def getInputDir(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχεία για μετονομασία")

        if dName == "":
            return

        self.inputDir.set(dName)
        self.ntrOutputDir.configure(state='readonly')
        self.btnOpenOutputDir.configure(state='normal')

    def getOutputDir(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία")

        if dName == "":
            return

        self.outputDir.set(dName)
        self.btnRun.configure(state='normal')

    def rename(self):
        inputDir = self.inputDir.get()
        outputDir = self.outputDir.get()

        for path, dirs, files in os.walk(inputDir):
            for file in files:
                subdir = path.replace(inputDir, "").lstrip("\\")
                destinationDir = os.path.join(outputDir, subdir)

                if not os.path.exists(destinationDir):
                    os.makedirs(destinationDir)

                target_name = file.replace(self.textToRemove.get(), "")
                copyfile(os.path.join(path, file), os.path.join(destinationDir, target_name))

    def run(self):
        if self.textToRemove.get() == '':
            showwarning(title='Μη συμπλήρωση πεδίου', message='Πρέπει να συμπληρώσετε το πεδίο "Κείμενο για αφαίρεση από το όνομα των αρχείων" για να εκτελέσετε τη μετονομασία των αρχείων.')
            return

        self.btnOpenInputDir.configure(state='disabled')
        self.btnOpenOutputDir.configure(state='disabled')
        self.btnRun.configure(state='disabled')

        self.rename()

        showinfo(title='Ολοκλήρωση Εκτέλεσης',
                 message='Η μετονομασία των αρχείων ολοκληρώθηκε.')

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lTextToRemove = Label(self.fData, text="Κείμενο για αφαίρεση από\nτο όνομα των αρχείων:")
        self.lTextToRemove.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.textToRemove = StringVar()
        self.textToRemove.set('')
        self.ntrTextToRemove = Entry(self.fData, width=128, textvariable=self.textToRemove)
        self.ntrTextToRemove.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.lInputDir = Label(self.fData, text="Φάκελος εισόδου:")
        self.lInputDir.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.inputDir = StringVar()
        self.inputDir.set('')
        self.ntrInputDir = Entry(self.fData, width=128, textvariable=self.inputDir)
        self.ntrInputDir.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenInputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getInputDir)
        self.btnOpenInputDir.grid(column=2, row=1, padx=10, pady=10)

        self.lOutputDir = Label(self.fData, text="Φάκελος εξόδου:")
        self.lOutputDir.grid(column=0, row=2, padx=10, pady=10, sticky=E)

        self.outputDir = StringVar()
        self.outputDir.set('')
        self.ntrOutputDir = Entry(self.fData, width=128, state='disabled', textvariable=self.outputDir)
        self.ntrOutputDir.grid(column=1, row=2, padx=10, pady=10, sticky=W)

        self.btnOpenOutputDir = Button(self.fData, text="Επιλέξτε φάκελο...", state='disabled',
                                       command=self.getOutputDir)
        self.btnOpenOutputDir.grid(column=2, row=2, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showinfo
from tkinter.ttk import *
from os import listdir
from os.path import isfile, join


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title("Δημιουργία λίστας με τα ονόματα των αρχείων που περιέχονται σε έναν φάκελο")
        self.window.resizable(False, False)
        self.create_widgets()

    def getFilesDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχεία")

        if dName == "":
            return

        self.filesDirName.set(dName)
        self.btnRun.configure(state='normal')

    def run(self):
        self.btnRun.configure(state='disabled')
        self.create_filenames_list()
        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η δημιουργία της λίστας ολοκληρώθηκε επιτυχώς.")
        self.window.destroy()

    def create_filenames_list(self):
        mypath = self.filesDirName.get()
        files = [f for f in listdir(mypath) if isfile(join(mypath, f))]

        with open('output.txt', 'w', encoding='utf-8') as f:
            for file in files:
                f.write(f'{file}\n')

        files_w_sc = list()
        for file in files:
            try:
                file.encode('windows-1253')
            except:
                files_w_sc.append(file)

        if len(files_w_sc) > 0:
            with open('special_chars.txt', 'w', encoding='utf-8') as f:
                for file in files_w_sc:
                    f.write(f'{file}\n')

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lData = Label(self.fData, text="Φάκελος με αρχεία:")
        self.lData.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.filesDirName = StringVar()
        self.ntrFilesDirName = Entry(self.fData, width=100, state='readonly', textvariable=self.filesDirName)
        self.ntrFilesDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenData = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.getFilesDirName)
        self.btnOpenData.grid(column=2, row=0, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

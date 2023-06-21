from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo, showwarning
import fitz
import os
import re
from shutil import copyfile


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title("Μετονομασία αρχείων με προσθήκη της ΑΔΑ")
        self.window.resizable(False, False)
        self.create_widgets()

    def get_output_dir_name(self):
        d_name = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που θα αποθηκευτούν "
                                                                     "τα αρχεία")

        if d_name == "":
            return

        self.outputDirName.set(d_name)
        self.btnRun.configure(state='normal')

    def get_input_dir_name(self):
        d_name = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχικά αρχεία")

        if d_name == "":
            return

        self.inputDirName.set(d_name)
        self.ntrOutputDirName.configure(state='readonly')
        self.btnOpenOutputDir.configure(state='normal')

    def verify_ada(self, ada):
        if len(ada) != 14:
            return False

        if ada[10] != '-':
            return False

        return True

    def find_ada(self, text):
        text_list = text.split('\n')

        ada = ''
        for i in range(len(text_list) - 1):
            if text_list[i].startswith('ΑΔΑ: ') and text_list[i + 1].startswith('Ministry of Digital'):
                ada = text_list[i].replace('ΑΔΑ: ', '')
                break

        return ada

    def parse_pdf_for_ada(self, filename):
        pdf = fitz.open(filename)
        page = pdf[0]
        text = page.get_text("text")
        pdf.close()

        ada = self.find_ada(text)

        if self.verify_ada(ada):
            return ada

        return None

    def run(self):
        self.btnRun.configure(state='disabled')

        input_dir = self.inputDirName.get()
        output_dir = self.outputDirName.get()

        showinfo(title="Έναρξη εκτέλεσης",
                 message="Mην τερματίσετε την εφαρμογή μέχρι να εμφανιστεί το μήνυμα της ολοκλήρωσης.")

        errors = ''
        err_count = 0
        for path, dirs, files in os.walk(input_dir):
            for file in files:
                if file.endswith(".pdf"):
                    ada = self.parse_pdf_for_ada(os.path.join(path, file))

                    if ada is None:
                        err_count += 1
                        errors += f"{err_count:2}: Στo αρχείο '{file}' δεν εντοπίστηκε ΑΔΑ.\n"
                        continue

                    dot_pos = file.rindex('.')
                    filename = re.sub(' +', ' ', file[:dot_pos])
                    ext = file[dot_pos:]

                    subdir = path.replace(input_dir, "").lstrip("\\")
                    destination_dir = os.path.join(output_dir, subdir)

                    if not os.path.exists(destination_dir):
                        os.makedirs(destination_dir)

                    target_name = f'{filename} [{ada}]{ext}'
                    copyfile(os.path.join(path, file), os.path.join(destination_dir, target_name))

        if errors == '':
            showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η μετονομασία των αρχείων ολοκληρώθηκε.")
        else:
            print(20 * '-', ' Αρχεία στα οποία δεν εντοπίστηκε ΑΔΑ ', 20 * '-')
            print(errors)
            showwarning(title="Ολοκλήρωση εκτέλεσης με λάθη", message="Η μετονομασία των αρχείων ολοκληρώθηκε. "
                                                                      "Υπάρχουν αρχεία στα οποία δεν εντοπίστηκε ΑΔΑ.")
        self.window.destroy()

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lInputDirName = Label(self.fData, text="Φάκελος με τα αρχικά αρχεία:")
        self.lInputDirName.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.inputDirName = StringVar()
        self.ntrInputDirName = Entry(self.fData, width=128, state='readonly', textvariable=self.inputDirName)
        self.ntrInputDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenInputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.get_input_dir_name)
        self.btnOpenInputDir.grid(column=2, row=0, padx=10, pady=10)

        self.lOutputDirName = Label(self.fData, text="Φάκελος για αποθήκευση των αρχείων:")
        self.lOutputDirName.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.outputDirName = StringVar()
        self.ntrOutputDirName = Entry(self.fData, width=128, state='disabled', textvariable=self.outputDirName)
        self.ntrOutputDirName.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenOutputDir = Button(self.fData, text="Επιλέξτε φάκελο...", command=self.get_output_dir_name,
                                       state='disabled')
        self.btnOpenOutputDir.grid(column=2, row=1, padx=10, pady=10)

        self.btnRun = Button(self.fData, text="Εκτέλεση", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()

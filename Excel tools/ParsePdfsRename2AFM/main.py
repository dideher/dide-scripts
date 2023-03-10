from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showinfo, showwarning
from tkinter.ttk import *
import os
from shutil import copyfile
import fitz
import re


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title("Μετονομασία αρχείων με ΑΦΜ")
        self.window.resizable(False, False)
        self.create_widgets()

    def get_input_dir(self):
        d_name = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχεία για μετονομασία")

        if d_name == "":
            return

        self.input_dir.set(d_name)
        self.ntr_output_dir.configure(state='readonly')
        self.btn_get_output_dir.configure(state='normal')

    def get_output_dir(self):
        d_name = filedialog.askdirectory(initialdir="./data/",
                                         title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία")

        if d_name == "":
            return

        self.output_dir.set(d_name)
        self.btn_run.configure(state='normal')

    def scan_file(self, filename):
        pdf = fitz.open(filename)
        page = pdf[0]
        page_text = page.get_text("text")
        pdf.close()

        return self.find_afm(page_text)

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

    def find_afm(self, text):
        text_list = re.sub(' +', ' ', text.replace('\n', ' ')).split(' ')

        afm_list = list()

        for item in text_list:
            if self.verify_afm(item) and item != '099709779':
                afm_list.append(item)

        if afm_list:
            return afm_list[-1]
        else:
            return None

    def rename_to_afm(self):
        input_dir = self.input_dir.get()
        output_dir = self.output_dir.get()

        showinfo(title="Έναρξη εκτέλεσης",
                 message="Mην τερματίσετε την εφαρμογή μέχρι να εμφανιστεί το μήνυμα της ολοκλήρωσης.")

        errors = ''
        err_count = 0
        for path, dirs, files in os.walk(input_dir):
            for file in files:
                if '.pdf' in file:
                    dotPos = file.rindex('.')
                    afm = self.scan_file(os.path.join(path, file))
                    if afm is None:
                        err_count += 1
                        errors += f"{err_count:2}: Στo αρχείο '{file}' δεν εντοπίστηκε ΑΦΜ.\n"
                        continue
                    ext = file[dotPos:]

                    subdir = path.replace(input_dir, "").lstrip("\\")
                    destination_dir = os.path.join(output_dir, subdir)

                    if not os.path.exists(destination_dir):
                        os.makedirs(destination_dir)

                    target_name = afm + ext
                    copyfile(os.path.join(path, file), os.path.join(destination_dir, target_name))

        if errors == '':
            showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η μετονομασία των αρχείων ολοκληρώθηκε.")
        else:
            print(20 * '-', ' Αρχεία στα οποία δεν εντοπίστηκε ΑΦΜ ', 20 * '-')
            print(errors)
            showwarning(title="Ολοκλήρωση εκτέλεσης με λάθη", message="Η μετονομασία των αρχείων ολοκληρώθηκε. "
                                                                      "Υπάρχουν αρχεία στα οποία δεν εντοπίστηκε ΑΦΜ.")

    def run(self):
        self.btn_get_input_dir.configure(state='disabled')
        self.btn_get_output_dir.configure(state='disabled')
        self.btn_run.configure(state='disabled')

        self.rename_to_afm()

    def create_widgets(self):
        self.f_data = Frame(self.window)

        self.l_input_dir = Label(self.f_data, text="Φάκελος εισόδου:")
        self.l_input_dir.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.input_dir = StringVar()
        self.input_dir.set('')
        self.ntr_input_dir = Entry(self.f_data, width=128, state='readonly', textvariable=self.input_dir)
        self.ntr_input_dir.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btn_get_input_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", command=self.get_input_dir)
        self.btn_get_input_dir.grid(column=2, row=1, padx=10, pady=10)

        self.l_output_dir = Label(self.f_data, text="Φάκελος εξόδου:")
        self.l_output_dir.grid(column=0, row=2, padx=10, pady=10, sticky=E)

        self.output_dir = StringVar()
        self.output_dir.set('')
        self.ntr_output_dir = Entry(self.f_data, width=128, state='disabled', textvariable=self.output_dir)
        self.ntr_output_dir.grid(column=1, row=2, padx=10, pady=10, sticky=W)

        self.btn_get_output_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", state='disabled',
                                         command=self.get_output_dir)
        self.btn_get_output_dir.grid(column=2, row=2, padx=10, pady=10)

        self.btn_run = Button(self.f_data, text="Εκτέλεση", command=self.run, state='disabled')
        self.btn_run.grid(column=1, row=10, padx=10, pady=10)

        self.f_data.pack()


gui = GUI()
gui.window.mainloop()

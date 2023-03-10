from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo
import os
from shutil import copyfile


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title('"Ισοπέδωση" ενός δέντρου αρχείων')
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

    def clean_spaces(self, text):
        while '  ' in text:
            text = text.replace('  ', ' ')

        return text

    def run(self):
        self.btnRun.configure(state='disabled')

        input_dir = self.inputDirName.get()
        output_dir = self.outputDirName.get()

        showinfo(title="Έναρξη εκτέλεσης",
                 message="Mην τερματίσετε την εφαρμογή μέχρι να εμφανιστεί το μήνυμα της ολοκλήρωσης.")

        for path, dirs, files in os.walk(input_dir):
            for file in files:
                dot_pos = file.rindex('.')
                filename = self.clean_spaces(file[:dot_pos])
                ext = file[dot_pos:]

                subdir = path.replace(input_dir, "").replace("\\", "_")
                destination_dir = output_dir

                if not os.path.exists(destination_dir):
                    os.makedirs(destination_dir)

                if len(subdir) > 0:
                    target_name = f'{subdir[1:]}_{filename}{ext}'
                else:
                    target_name = f'{filename}{ext}'
                copyfile(os.path.join(path, file), os.path.join(destination_dir, target_name))

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η μετονομασία και αντιγραφή των αρχείων στον φάκελο "
                                                       "προορισμού ολοκληρώθηκε.")
        self.window.destroy()

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lInputDirName = Label(self.fData, text="Φάκελος με τα αρχικά αρχεία:")
        self.lInputDirName.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.inputDirName = StringVar()
        self.ntrInputDirName = Entry(self.fData, width=128, state='readonly', textvariable=self.inputDirName)
        self.ntrInputDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenData = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.get_input_dir_name)
        self.btnOpenData.grid(column=2, row=0, padx=10, pady=10)

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

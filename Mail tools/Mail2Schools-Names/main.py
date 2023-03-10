from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from tkinter.messagebox import showwarning, showinfo
from tkinter.ttk import *
from openpyxl import *
from os import walk
import yagmail
from threading import Thread


class GUI:
    def __init__(self):
        self.window = Tk()
        self.window.bind_all("<Key>", self._onKeyRelease, "+")
        self.window.title("Αποστολή μαζικής αλληλογραφίας σε Σχολικές Μονάδες")
        self.window.resizable(False, False)
        self.create_widgets()

    def _onKeyRelease(self, event):
        ctrl = (event.state & 0x4) != 0
        if event.keycode == 88 and ctrl and event.keysym.lower() != "x":
            event.widget.event_generate("<<Cut>>")

        if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
            event.widget.event_generate("<<Paste>>")

        if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
            event.widget.event_generate("<<Copy>>")

    def debugCkbChange(self):
        if self.debugChecked.get():
            self.ntrDebugRecipient.configure(state='normal')
        else:
            self.ntrDebugRecipient.configure(state='disabled')

    def parseXlsxData(self, test):
        workbook = load_workbook(filename=self.addressBookFilename.get())
        sheet = workbook.active

        self.schools = list()

        for row in sheet.iter_rows():
            entry = list()
            for cell in row:
                if cell.value is None:
                    text = ""
                else:
                    text = str(cell.value).strip()

                entry.append(text)

            self.schools.append(entry)

        sender_email = self.gmail.get()
        password = self.passwd.get()
        inputDirectory = self.filesDirName.get()
        debugMode = self.debugChecked.get()

        self.stPreview.delete('1.0', END)

        previewText = ""
        if test:
            if debugMode:
                previewText += "-" * 20 + " Δοκιμαστική λειτουργία " + "-" * 20 + "\n"
                previewText += "Αποστολή στον παραλήπτη: " + self.debugRecipient.get() + "\n"
                previewText += 64 * "-" + "\n\n"
            previewText += "-" * 20 + " Θέμα " + "-" * 20 + "\n"
            previewText += self.subject.get() + "\n\n"
            previewText += "-" * 20 + " Σώμα " + "-" * 20 + "\n"
            previewText += self.stBody.get('1.0', END) + "\n\n"
            previewText += "-" * 20 + " Κοινό συνημμένο " + "-" * 20 + "\n"
            if self.attachmentFilename.get() == "":
                previewText += "Χωρίς κοινό συνημμένο." + "\n\n"
            else:
                previewText += self.attachmentFilename.get() + "\n\n"
            previewText += "-" * 20 + " Βιβλίο διευθύνσεων " + "-" * 20 + "\n"
            previewText += self.addressBookFilename.get() + "\n\n"
            previewText += "-" * 20 + " Φάκελος αρχείων " + "-" * 20 + "\n"
            previewText += self.filesDirName.get() + "\n\n"
            previewText += "-" * 20 + " Αποστολή ... " + "-" * 20 + "\n\n"

        excludedSchools = ""
        errors = ""
        exceptions = ""

        if not test:
            showinfo(title="Έναρξη αποστολής",
                     message="Μην τερματίσετε την εφαρμογή μέχρι να εμφανιστεί το μήνυμα ολοκλήρωσης της αποστολής.")
            yag = yagmail.SMTP(sender_email, password=password)

        filesCounter = 0
        sendCounter = 0
        for sch in self.schools[1:]:
            schoolName = sch[0]
            schoolCode = sch[1]
            excluded = sch[3]

            if not debugMode:
                recipient_email = sch[2]
            else:
                recipient_email = self.debugRecipient.get()

            if not excluded:
                attachmentsList = list()
                appendCommonAttachment = True

                for path, dirs, files in walk(inputDirectory):
                    for file in files:
                        spaceBeforeName = " " + schoolName
                        dashBeforeName = "-" + schoolName

                        if schoolName == file[:len(schoolName)] or spaceBeforeName in file or dashBeforeName in file:
                            filesCounter += 1

                            if self.attachmentFilename.get() != "" and appendCommonAttachment:
                                attachmentsList.append(self.attachmentFilename.get())
                                appendCommonAttachment = False

                            attachmentsList.append(path + "\\" + file)

                if len(attachmentsList) == 0:
                    errors += f"To σχολείο '{schoolName} ({schoolCode})' δεν έχει αρχεία για αποστολή.\n"
                else:
                    if test:
                        previewText += "-" * 20 + f" {schoolName} ({schoolCode}) " + "-" * 20

                        if len(attachmentsList) == 1:
                            previewText += f"\nΑποστολή του παρκάτω αρχείου στον παραλήπτη '{recipient_email}':\n"
                            previewText += f"1: '{attachmentsList[0]}'\n"
                        else:
                            previewText += f"\nΑποστολή των παρακάτω αρχείων στον παραλήπτη '{recipient_email}':\n"
                            i = 0
                            skipAttachment = True
                            for a in attachmentsList:
                                if self.attachmentFilename.get() != "" and skipAttachment:
                                    previewText += f"Κοινό συνημμένο: '{a}'\n"
                                    skipAttachment = False
                                else:
                                    i += 1
                                    previewText += f"{i}: '{a}'\n"
                    else:
                        try:
                            yag.send(to=recipient_email, subject=self.subject.get(),
                                     contents=self.stBody.get('1.0', END), attachments=attachmentsList)
                        except:
                            exceptions += f"Η αποστολή στον παραλήπτη '{recipient_email}' απέτυχε.\n"
                        else:
                            sendCounter += 1
                            self.stPreview.insert(INSERT,
                                                  f"{sendCounter:2d}) Η αποστολή στον παραλήπτη '{schoolName}: {recipient_email}' πέτυχε.\n")
            else:
                excludedSchools += f"Για το σχολείο '{schoolName}' υπήρχε εξαίρεση.\n"

        if test:
            previewText += "\n" + 64 * "-" + "\n"
            previewText += f"Πλήθος αρχείων: {filesCounter} \n"
            previewText += 64 * "-" + "\n\n"

            if excludedSchools != "":
                previewText += "\n" + "-" * 20 + " Εξαιρέσεις " + "-" * 20 + "\n"
                previewText += excludedSchools

            if errors != "":
                previewText += "\n" + "-" * 20 + " Λάθη " + "-" * 20 + "\n"
                previewText += errors
        else:
            if exceptions != "":
                previewText += "\n" + "-" * 20 + " Αποτυχημένες αποστολές " + "-" * 20 + "\n"
                previewText += exceptions

        self.stPreview.insert(INSERT, previewText)

        if not test:
            showinfo(title="Ολοκλήρωση αποστολής", message="Η αποστολή ολοκληρώθηκε.")

    def getAttachmentFilename(self):
        fName = filedialog.askopenfilename(initialdir="./data/", title="Επιλέξτε το έγγραφο",
                                           filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

        if fName == "":
            return

        self.attachmentFilename.set(fName)

    def nextGoToTabSettings(self):
        if self.subject.get() == "" and self.stBody.get('1.0', END) == "\n":
            showwarning(title="Κενά πεδία...", message="Παρακαλώ συμπληρώστε τα πεδία 'Θέμα' και 'Σώμα' του μηνύματος.")
            return
        elif self.subject.get() == "":
            showwarning(title="Κενό πεδίο...", message="Παρακαλώ συμπληρώστε το πεδίο 'Θέμα' του μηνύματος.")
            return
        elif self.stBody.get('1.0', END) == "\n":
            showwarning(title="Κενό πεδίο...", message="Παρακαλώ συμπληρώστε το πεδίο 'Σώμα' του μηνύματος.")
            return

        self.tabControl.tab(0, state='disabled')
        self.tabControl.tab(1, state='normal')
        self.tabControl.select(1)

    def nextGoToTabPreview(self):
        if self.addressBookFilename.get() == "" or self.filesDirName.get() == "" or self.gmail.get() == "" or self.passwd.get() == "":
            showwarning(title="Κενά πεδία...", message="Παρακαλώ συμπληρώστε όλα τα πεδία για να προχωρήσετε.")
            return

        if self.debugChecked.get() and self.debugRecipient.get() == "":
            showwarning(title="Κενό πεδίο...",
                        message="Έχετε ενεργοποιήσει τη δοκιμαστική λειτουργία. Για να προχωρήσετε πρέπει υποχρεωτικά να συμπληρώσετε το πεδίο 'Παραλήπτης για δοκιμαστική αποστολή'.")
            return

        self.tabControl.tab(1, state='disabled')
        self.tabControl.tab(2, state='normal')
        self.tabControl.select(2)

    def getAddressBookFilename(self):
        fName = filedialog.askopenfilename(initialdir="./mail/", title="Επιλέξτε το Βιβλίο Διευθύνσεων",
                                           filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

        if fName == "":
            return

        self.addressBookFilename.set(fName)

    def getFilesDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχεία προς αποστολή")

        if dName == "":
            return

        self.filesDirName.set(dName)

    def prevGoToTabSubjectBody(self):
        self.tabControl.tab(1, state='disabled')
        self.tabControl.tab(0, state='normal')
        self.tabControl.select(0)

    def prevGoToTabSettings(self):
        self.stPreview.delete('1.0', END)
        self.notPreviewed = True
        self.btnSend.configure(state='disabled')

        self.tabControl.tab(2, state='disabled')
        self.tabControl.tab(1, state='normal')
        self.tabControl.select(1)

    def preview(self):
        self.notPreviewed = False
        self.btnSend.configure(state='normal')
        self.parseXlsxData(test=True)

    def send(self):
        self.btnPrevGoToSettings.configure(state='disabled')
        self.btnPreview.configure(state='disabled')
        self.btnSend.configure(state='disabled')

        self.runThread = Thread(target=self.parseXlsxData, args=[False])
        self.runThread.start()

    def create_widgets(self):
        # Tabs
        self.tabControl = Notebook(self.window)
        self.tabSubjectBody = Frame(self.tabControl)
        self.tabControl.add(self.tabSubjectBody, text="Δημιουργία μηνύματος")
        self.tabSettings = Frame(self.tabControl)
        self.tabControl.add(self.tabSettings, text="Επιλογές/Ρυθμίσεις", state='disabled')
        self.tabPreview = Frame(self.tabControl)
        self.tabControl.add(self.tabPreview, text="Προεπισκόπηση/Αποστολή", state='disabled')
        self.tabControl.pack(expand=1, fill="both")

        # Tab: SubjectBody
        self.fSubjectBody = Frame(self.tabSubjectBody)

        self.lSubject = Label(self.fSubjectBody, text="Θέμα:")
        self.lSubject.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.subject = StringVar()
        self.ntrSubject = Entry(self.fSubjectBody, width=110, textvariable=self.subject)
        self.ntrSubject.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.lBody = Label(self.fSubjectBody, text="Σώμα:")
        self.lBody.grid(column=0, row=1, padx=10, pady=10, sticky=NE)

        self.stBody = ScrolledText(self.fSubjectBody, height=10, wrap=WORD)
        self.stBody.grid(column=1, row=1, padx=10, pady=10)

        self.lAttachment = Label(self.fSubjectBody, text="Κοινό συνημμένο αρχείο\n(π.χ. διαβιβαστικό):")
        self.lAttachment.grid(column=0, row=2, padx=10, pady=10, sticky=E)

        self.attachmentFilename = StringVar()
        self.ntrAttachmentFilename = Entry(self.fSubjectBody, width=110, state='readonly',
                                           textvariable=self.attachmentFilename)
        self.ntrAttachmentFilename.grid(column=1, row=2, padx=10, pady=10, sticky=W)

        self.btnOpenAttachment = Button(self.fSubjectBody, text="Επιλέξτε αρχείο...",
                                        command=self.getAttachmentFilename)
        self.btnOpenAttachment.grid(column=2, row=2, padx=10, pady=10)

        self.btnNextGoToSettings = Button(self.fSubjectBody, text="Επόμενο", command=self.nextGoToTabSettings)
        self.btnNextGoToSettings.grid(column=2, row=10, padx=10, pady=10)

        self.fSubjectBody.pack()

        # Tab: Settings
        self.fSettings = Frame(self.tabSettings)

        self.lAddressBook = Label(self.fSettings, text="Βιβλίο Διευθύνσεων:")
        self.lAddressBook.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.addressBookFilename = StringVar()
        self.ntrAddressBookFilename = Entry(self.fSettings, width=110, state='readonly',
                                            textvariable=self.addressBookFilename)
        self.ntrAddressBookFilename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenAddressBook = Button(self.fSettings, text="Επιλέξτε αρχείο...", command=self.getAddressBookFilename)
        self.btnOpenAddressBook.grid(column=2, row=0, padx=10, pady=10)

        self.lFiles = Label(self.fSettings, text="Φάκελος με αρχεία προς αποστολή:")
        self.lFiles.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.filesDirName = StringVar()
        self.ntrFilesDirName = Entry(self.fSettings, width=110, state='readonly', textvariable=self.filesDirName)
        self.ntrFilesDirName.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnOpenFilesDirName = Button(self.fSettings, text="Επιλέξτε φάκελο...", command=self.getFilesDirName)
        self.btnOpenFilesDirName.grid(column=2, row=1, padx=10, pady=10)

        self.lGmail = Label(self.fSettings, text="Gmail:")
        self.lGmail.grid(column=0, row=2, padx=10, pady=10, sticky=E)

        self.gmail = StringVar()
        self.ntrGmail = Entry(self.fSettings, width=110, textvariable=self.gmail)
        self.ntrGmail.grid(column=1, row=2, padx=10, pady=10, sticky=W)

        self.lPasswd = Label(self.fSettings, text="Password:")
        self.lPasswd.grid(column=0, row=3, padx=10, pady=10, sticky=E)

        self.passwd = StringVar()
        self.ntrPasswd = Entry(self.fSettings, width=110, textvariable=self.passwd)
        self.ntrPasswd.grid(column=1, row=3, padx=10, pady=10, sticky=W)

        self.lfDebug = LabelFrame(self.fSettings, text="Δοκιμαστική λειτουργία")
        self.lfDebug.grid(column=0, row=4, columnspan=3, padx=10, pady=10)

        self.lDebugRecipient = Label(self.lfDebug, text="Παραλήπτης για δοκιμαστική\nαποστολή:")
        self.lDebugRecipient.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.debugRecipient = StringVar()
        self.ntrDebugRecipient = Entry(self.lfDebug, width=110, textvariable=self.debugRecipient, state='disabled')
        self.ntrDebugRecipient.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.debugChecked = BooleanVar()
        self.ckbDebugChecked = Checkbutton(self.lfDebug, text="Ενεργοποίηση δοκιμαστικής\nλειτουργίας",
                                           variable=self.debugChecked, command=self.debugCkbChange)
        self.ckbDebugChecked.grid(column=2, row=0, padx=10, pady=10)

        self.btnPrevGoToSubjectBody = Button(self.fSettings, text="Προηγούμενο", command=self.prevGoToTabSubjectBody)
        self.btnPrevGoToSubjectBody.grid(column=0, row=10, padx=10, pady=10)

        self.btnNextGoToPreview = Button(self.fSettings, text="Επόμενο", command=self.nextGoToTabPreview)
        self.btnNextGoToPreview.grid(column=2, row=10, padx=10, pady=10)

        self.fSettings.pack()

        # Tab: Preview
        self.fPreview = Frame(self.tabPreview)

        self.stPreview = ScrolledText(self.fPreview, width=120, height=16)
        self.stPreview.grid(column=0, row=0, columnspan=3, padx=10, pady=10)

        self.btnPrevGoToSettings = Button(self.fPreview, text="Προηγούμενο", command=self.prevGoToTabSettings)
        self.btnPrevGoToSettings.grid(column=0, row=10, padx=10, pady=10)

        self.btnPreview = Button(self.fPreview, text="Προεπισκόπηση", command=self.preview)
        self.btnPreview.grid(column=1, row=10, padx=10, pady=10)

        self.notPreviewed = True

        self.btnSend = Button(self.fPreview, text="Αποστολή", command=self.send, state='disabled')
        self.btnSend.grid(column=2, row=10, padx=10, pady=10)

        self.fPreview.pack()


gui = GUI()
gui.window.mainloop()

from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from tkinter.messagebox import showwarning, showerror, askquestion
from tkinter.ttk import *
import json
import opendata
import requests
import os
from datetime import datetime
from pathlib import Path
from threading import Thread


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Μαζική μεταφόρτωση Συμβάσεων στη Διαύγεια")
        self.window.resizable(False, False)
        self.create_widgets()
        self.configIsLoaded = self.loadConfig()


    def debugCkbChange(self):
        if self.debugChecked.get():
            self.login.set('10599_api')
            self.passwd.set('User@10599')
            self.metadataFilename.set('./json/SampleDecisionMetadata.json')
        else:
            if self.configIsLoaded:
                self.login.set(self.config['login'])
                self.passwd.set(self.config['password'])
                self.metadataFilename.set(self.config['jsonFile'])
            else:
                self.login.set('')
                self.passwd.set('')
                self.metadataFilename.set('')


    def publishDicisions(self):
        inputDir = self.filesDirName.get()
        outputDir = os.path.join(inputDir, "ADAs")
        if self.publish.get():
            Path(outputDir).mkdir(parents=True, exist_ok=True)

        for entry in os.scandir(inputDir):
            if entry.is_file():
                file = entry.name
                if file[-4:] == ".pdf":
                    self.metadataTemplate['subject'] = self.subject.get().replace('...', file[:-4])
                    now = datetime.now()
                    self.metadataTemplate['issueDate'] = now.strftime("%Y-%m-%dT%H:%M:%S.000Z")
                    self.metadataTemplate['protocolNumber'] = self.arProt.get()
                    self.publishDecision(file, inputDir, outputDir)

        if (self.counter == 0):
            self.stPreview.insert(INSERT, "-" * 20 + " ΔΕΝ ΕΓΙΝΕ ΜΕΤΑΦΟΡΤΩΣΗ ΑΡΧΕΙΩΝ " + "-" * 20 + "\n")
        else:
            self.stPreview.insert(INSERT, "-" * 20 + " ΟΛΟΚΛΗΡΩΣΗ ΜΕΤΑΦΟΡΤΩΣΕΩΝ " + "-" * 20 + "\n")

        if self.errors != "":
            self.stPreview.insert(INSERT, "-" * 20 + " ΛΑΘΗ " + "-" * 20 + "\n")
            self.stPreview.insert(INSERT, self.errors)

        self.btnOpenFilesDirName.configure(state='normal')
        self.ntrArProt.configure(state='normal')
        self.btnUpload.configure(state='normal')


    def upload(self):
        if not self.notUploaded:
            if askquestion('Προσοχή...', 'Έχετε ήδη μεταφορτώσει τα αρχεία αυτού του φακέλου. Θέλετε να επαναλάβετε τη διαδικασία με τον ίδιο φάκελο;', icon='warning') == 'no':
                self.btnUpload.configure(state='disabled')
                return

        if self.arProt.get() == "":
            showwarning(title="Κενό πεδίο...", message="Παρακαλώ συμπληρώστε τον Αρ. πρωτοκόλλου.")
            return

        self.notUploaded = False
        self.counter = 0
        self.errCounter = 0
        self.errors = ""
        self.stPreview.delete('1.0', END)

        self.btnOpenFilesDirName.configure(state='disabled')
        self.ntrArProt.configure(state='disabled')
        self.btnUpload.configure(state='disabled')
        self.runThread = Thread(target=self.publishDicisions)
        self.runThread.start()


    def nextGoToTabMetadata(self):
        if self.login.get() == "" or self.passwd.get() == "":
            showwarning(title="Κενά πεδία...", message="Παρακαλώ συμπληρώστε και τα δύο πεδία.")
            return

        if self.debugChecked.get():
            self.client = opendata.OpendataClient()
        else:
            self.client = opendata.OpendataClient('https://diavgeia.gov.gr/opendata')

        if self.metadataFilename.get() != '' and self.loadMetadata():
            self.updateMetadataInfo()

        self.client.set_credentials(self.login.get(), self.passwd.get())

        self.ckbDebugChecked.configure(state='disabled')
        self.ntrLogin.configure(state='disabled')
        self.ntrPasswd.configure(state='disabled')
        self.btnNextGoToMetadata.configure(state='disabled')
        self.tabControl.tab(1, state='normal')
        self.tabControl.select(1)


    def nextGoToTabFiles(self):
        self.btnOpenMetadata.configure(state='disabled')
        self.rbSubmit.configure(state='disabled')
        self.rbPublish.configure(state='disabled')
        self.metadataTemplate['publish'] = self.publish.get()
        self.btnNextGoToFiles.configure(state='disabled')
        self.tabControl.tab(2, state='normal')
        self.tabControl.select(2)


    def getFilesDirName(self):
        dName = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο με τα αρχεία προς μεταφόρτωση")

        if dName == "":
            return

        if dName != self.filesDirName.get():
            self.notUploaded = True

        self.filesDirName.set(dName)
        self.btnUpload.configure(state='normal')


    def loadConfig(self):
        try:
            json_file = open("./json/_config.json", 'r', encoding='utf-8')
        except:
            return False
        else:
            self.config = json.load(json_file)
            self.login.set(self.config['login'])
            self.passwd.set(self.config['password'])
            self.metadataFilename.set(self.config['jsonFile'])
            json_file.close()
            return True


    def loadMetadata(self):
        # Decision metadata
        try:
            json_file = open(self.metadataFilename.get(), 'r', encoding='utf-8')
        except:
            showerror("Εύρεση αρχείου...", "Δεν μπορεί να εντοπιστεί το αρχείο '{}'.".format(self.metadataFilename.get()))
            self.metadataFilename.set('')
            return False
        else:
            self.metadataTemplate = json.load(json_file)
            json_file.close()
            self.btnOpenMetadata.configure(state='disabled')
            return True


    def saveADAfile(self, ada, url, file, outputDir):
        r = requests.get(url, allow_redirects=True)

        outputFile = file.replace(".pdf", " {}.pdf".format(ada))
        open(os.path.join(outputDir, outputFile), 'wb').write(r.content)


    def publishDecision(self, file, inputDir, outputDir):
        # Decision document
        pdf_file = open(os.path.join(inputDir, file), 'rb')

        # Send request
        response = self.client.submit_decision(self.metadataTemplate, pdf_file)
        pdf_file.close()

        if response.status_code == 200:
            decision = response.json()
            ada = decision['ada']
            url = decision['documentUrl']

            self.counter += 1
            self.stPreview.insert(INSERT, "{:2d}) Η μεταφόρτωση του αρχείου '{}' ολοκληρώθηκε επιτυχώς. [Θέμα: {}, ΑΔΑ: {}]\n".format(self.counter, file, self.metadataTemplate['subject'], ada))
            if self.publish.get():
                self.saveADAfile(ada, url, file, outputDir)
        elif response.status_code == 400:
            self.errCounter += 1
            self.errors += "{:2d}) Η μεταφόρτωση του αρχείου '{}' απέτυχε. Σφάλμα στην υποβολή της πράξης.\n".format(self.errCounter, file)
            err_json = response.json()
            for err in err_json['errors']:
                self.errors += "{0}: {1}\n".format(err['errorCode'], err['errorMessage'])
        elif response.status_code == 401:
            self.errCounter += 1
            self.errors += "{:2d}) Η μεταφόρτωση του αρχείου '{}' απέτυχε. Σφάλμα αυθεντικοποίησης.\n".format(self.errCounter, file)
        elif response.status_code == 403:
            self.errCounter += 1
            self.errors += "{:2d}) Η μεταφόρτωση του αρχείου '{}' απέτυχε. Απαγόρευση πρόσβασης.\n".format(self.errCounter, file)
        else:
            self.errCounter += 1
            self.errors += "{:2d}) Η μεταφόρτωση του αρχείου '{}' απέτυχε. ERROR {}.\n".format(self.errCounter, file, str(response.status_code))


    def updateMetadataInfo(self):
        self.decisionType.set(self.client.get_decision_type(self.metadataTemplate['decisionTypeId'])['label'])
        self.thk.set(self.client.get_THK_type_label((self.metadataTemplate['thematicCategoryIds'])[0]))
        self.organization.set(self.client.get_organization(self.metadataTemplate['organizationId'])['label'])
        self.unit.set(self.client.get_unit(self.metadataTemplate['unitIds'][0])['label'])

        firstName = self.client.get_signer(self.metadataTemplate['signerIds'][0])['firstName']
        lastName = self.client.get_signer(self.metadataTemplate['signerIds'][0])['lastName']
        self.signer.set("{} {}".format(firstName, lastName))

        self.subject.set(self.metadataTemplate['subject'])
        self.contractType.set(self.metadataTemplate['extraFieldValues']['contractType'])
        if self.metadataTemplate['extraFieldValues']['financedProject']:
            self.financedProject.set('ΝΑΙ')
        else:
            self.financedProject.set('ΟΧΙ')

        self.arProt.set(self.metadataTemplate['protocolNumber'])
        self.btnNextGoToFiles.configure(state='normal')


    def getMetadataFilename(self):
        fName = filedialog.askopenfilename(initialdir="./json/", title="Επιλέξτε το αρχείο μεταδεδομένων",
                                           filetypes=(("json files", "*.json"), ("all files", "*.*")))

        if fName == "":
            return

        self.metadataFilename.set(fName)
        if self.loadMetadata():
            self.updateMetadataInfo()


    def create_widgets(self):
        # Tabs
        self.tabControl = Notebook(self.window)
        self.tabLogin = Frame(self.tabControl)
        self.tabControl.add(self.tabLogin, text="Σύνδεση")
        self.tabMetadata = Frame(self.tabControl)
        self.tabControl.add(self.tabMetadata, text="Μεταδεδομένα", state='disabled')
        self.tabFiles = Frame(self.tabControl)
        self.tabControl.add(self.tabFiles, text="Αρχεία", state='disabled')
        self.tabControl.pack(expand=1, fill="both")

        # Tab: Login
        self.fLogin = Frame(self.tabLogin)

        self.lLogin = Label(self.fLogin, text="Όνομα Χρήστη:")
        self.lLogin.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.login = StringVar()
        self.ntrLogin = Entry(self.fLogin, width=80, textvariable=self.login)
        self.ntrLogin.grid(column=1, row=0, columnspan=2, padx=10, pady=10, sticky=W)

        self.lPasswd = Label(self.fLogin, text="Κωδικός:")
        self.lPasswd.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.passwd = StringVar()
        self.ntrPasswd = Entry(self.fLogin, width=80, textvariable=self.passwd)
        self.ntrPasswd.grid(column=1, row=1, columnspan=2, padx=10, pady=10, sticky=W)

        self.debugChecked = BooleanVar()
        self.debugChecked.set(False)
        self.ckbDebugChecked = Checkbutton(self.fLogin, text="Ενεργοποίηση δοκιμαστικής λειτουργίας", variable=self.debugChecked, command=self.debugCkbChange)
        self.ckbDebugChecked.grid(column=1, row=2, padx=10, pady=10)

        self.btnNextGoToMetadata = Button(self.fLogin, text="Επόμενο", command=self.nextGoToTabMetadata)
        self.btnNextGoToMetadata.grid(column=1, row=10, padx=10, pady=10)

        self.fLogin.pack()

        # Tab: Metadata
        self.fMetadata = Frame(self.tabMetadata)

        self.lMetadataFile = Label(self.fMetadata, text="Αρχείο:")
        self.lMetadataFile.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.metadataFilename = StringVar()
        self.metadataFilename.set('')
        self.ntrMetadataFilename = Entry(self.fMetadata, width=128, state='readonly', textvariable=self.metadataFilename)
        self.ntrMetadataFilename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenMetadata = Button(self.fMetadata, text="Επιλέξτε αρχείο...", command=self.getMetadataFilename)
        self.btnOpenMetadata.grid(column=2, row=0, padx=10, pady=10)

        self.lfVerifyFrame = LabelFrame(self.fMetadata, text="Μεταδεδομένα")
        self.lfVerifyFrame.grid(column=0, row=1, columnspan=3, padx=10, pady=10)

        self.lDecisionType = Label(self.lfVerifyFrame, text="Είδος:")
        self.lDecisionType.grid(column=0, row=0, padx=10, pady=5, sticky=E)
        self.decisionType = StringVar()
        self.ntrDecisionType = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.decisionType)
        self.ntrDecisionType.grid(column=1, row=0, columnspan=2, padx=10, pady=5, sticky=W)

        self.lTHK = Label(self.lfVerifyFrame, text="Θεματικές κατηγορίες:")
        self.lTHK.grid(column=0, row=1, padx=10, pady=5, sticky=E)
        self.thk = StringVar()
        self.ntrTHK = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.thk)
        self.ntrTHK.grid(column=1, row=1, columnspan=2, padx=10, pady=5, sticky=W)

        self.lOrganization = Label(self.lfVerifyFrame, text="Φορέας:")
        self.lOrganization.grid(column=0, row=2, padx=10, pady=5, sticky=E)
        self.organization = StringVar()
        self.ntrOrganization = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.organization)
        self.ntrOrganization.grid(column=1, row=2, columnspan=2, padx=10, pady=5, sticky=W)

        self.lUnit = Label(self.lfVerifyFrame, text="Μονάδα:")
        self.lUnit.grid(column=0, row=3, padx=10, pady=5, sticky=E)
        self.unit = StringVar()
        self.ntrUnit = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.unit)
        self.ntrUnit.grid(column=1, row=3, columnspan=2, padx=10, pady=5, sticky=W)

        self.lSigner = Label(self.lfVerifyFrame, text="Υπογράφων:")
        self.lSigner.grid(column=0, row=4, padx=10, pady=5, sticky=E)
        self.signer = StringVar()
        self.ntrSigner = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.signer)
        self.ntrSigner.grid(column=1, row=4, columnspan=2, padx=10, pady=5, sticky=W)

        self.lSubject = Label(self.lfVerifyFrame, text="Θέμα:")
        self.lSubject.grid(column=0, row=5, padx=10, pady=5, sticky=E)
        self.subject = StringVar()
        self.ntrSubject = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.subject)
        self.ntrSubject.grid(column=1, row=5, columnspan=2, padx=10, pady=5, sticky=W)

        self.lContractType = Label(self.lfVerifyFrame, text="Είδος πράξης:")
        self.lContractType.grid(column=0, row=6, padx=10, pady=5, sticky=E)
        self.contractType = StringVar()
        self.ntrContractType = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.contractType)
        self.ntrContractType.grid(column=1, row=6, columnspan=2, padx=10, pady=5, sticky=W)

        self.lFinancedProject = Label(self.lfVerifyFrame, text="Συγχρηματοδοτούμενο έργο:")
        self.lFinancedProject.grid(column=0, row=7, padx=10, pady=5, sticky=E)
        self.financedProject = StringVar()
        self.ntrFinancedProject = Entry(self.lfVerifyFrame, width=110, state='readonly', textvariable=self.financedProject)
        self.ntrFinancedProject.grid(column=1, row=7, columnspan=2, padx=10, pady=5, sticky=W)

        self.lfPublishFrame = LabelFrame(self.fMetadata, text="Υποβολή / Ανάρτηση")
        self.lfPublishFrame.grid(column=1, row=2, padx=10, pady=10)

        self.publish = BooleanVar()
        self.rbSubmit = Radiobutton(self.lfPublishFrame, text="Υποβολή", value=False, variable=self.publish)
        self.rbSubmit.grid(column=0, row=0)
        Label(self.lfPublishFrame, text="                ").grid(column=1, row=0)
        self.rbPublish = Radiobutton(self.lfPublishFrame, text="Ανάρτηση", value=True, variable=self.publish)
        self.rbPublish.grid(column=2, row=0)
        self.publish.set(True)

        self.btnNextGoToFiles = Button(self.fMetadata, text="Επόμενο", command=self.nextGoToTabFiles, state='disabled')
        self.btnNextGoToFiles.grid(column=2, row=10, padx=10, pady=10)

        self.fMetadata.pack()

        # Tab: Files
        self.fFiles = Frame(self.tabFiles)

        self.lFiles = Label(self.fFiles, text="Φάκελος με αρχεία προς μεταφόρτωση:")
        self.lFiles.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.filesDirName = StringVar()
        self.ntrFilesDirName = Entry(self.fFiles, width=110, state='readonly', textvariable=self.filesDirName)
        self.ntrFilesDirName.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenFilesDirName = Button(self.fFiles, text="Επιλέξτε φάκελο...", command=self.getFilesDirName)
        self.btnOpenFilesDirName.grid(column=2, row=0, padx=10, pady=10)

        self.lArProt = Label(self.fFiles, text="Αρ. πρωτοκόλλου:")
        self.lArProt.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.arProt = StringVar()
        self.ntrArProt = Entry(self.fFiles, width=110, textvariable=self.arProt)
        self.ntrArProt.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btnUpload = Button(self.fFiles, text="Μεταφόρτωση", command=self.upload, state='disabled')
        self.btnUpload.grid(column=1, row=2, padx=10, pady=10)

        self.stPreview = ScrolledText(self.fFiles, width=120, height=16)
        self.stPreview.grid(column=0, row=3, columnspan=3, padx=10, pady=10)

        self.fFiles.pack()


gui = GUI()
gui.window.mainloop()

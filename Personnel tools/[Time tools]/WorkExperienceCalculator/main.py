from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.messagebox import showwarning


class GUI():
    def __init__(self):
        self.normalWE = list()
        self.hourlyPaidWE = list()
        self.nCounter = 0
        self.hpCounter = 0
        self.window = Tk()

        self.window.title("Υπολογισμός Προϋπηρεσίας")
        self.window.resizable(False, False)
        self.create_widgets()


    def addWE(self):
        try:
            y1 = int(self.year.get())
            m1 = int(self.month.get())
            d1 = int(self.day.get())

            if y1 < 0 or m1 < 0 or d1 < 0:
                showwarning("Προσοχή ...", "Όλα τα πεδία πρέπει να είναι συμπληρωμένα με θετικούς αριθμούς.")
                return
        except:
            showwarning("Προσοχή ...", "Όλα τα πεδία πρέπει να είναι συμπληρωμένα με αριθμούς.")
        else:
            if self.hourlyPaid.get():
                self.hpCounter += 1
                self.stHourlyPaidWE.configure(state='normal')
                self.stHourlyPaidWE.insert(INSERT, "{:2d}) {:02d}-{:02d}-{:02d}\n".format(self.hpCounter, y1, m1, d1))
                self.stHourlyPaidWE.configure(state='disabled')
                self.hourlyPaidWE.append([y1, m1, d1])
            else:
                self.nCounter += 1
                self.stNormalWE.configure(state='normal')
                self.stNormalWE.insert(INSERT, "{:2d}) {:02d}-{:02d}-{:02d}\n".format(self.nCounter, y1, m1, d1))
                self.stNormalWE.configure(state='disabled')
                self.normalWE.append([y1, m1, d1])
        finally:
            self.clean()


    def calcTotal(self, weList, daysOfMonth):
        sumY = 0
        sumM = 0
        sumD = 0
        for entry in weList:
            sumY += entry[0]
            sumM += entry[1]
            sumD += entry[2]

        totalD = sumD % daysOfMonth
        totalM = (sumM + sumD // daysOfMonth) % 12
        totalY = sumY + (sumM + sumD // daysOfMonth) // 12

        return [totalY, totalM, totalD]


    def clean(self):
        self.year.set('')
        self.month.set('')
        self.day.set('')

        self.yearsTNWE.set('')
        self.monthsTNWE.set('')
        self.daysTNWE.set('')

        self.yearsTHPWE.set('')
        self.monthsTHPWE.set('')
        self.daysTHPWE.set('')

        self.yearsTWE.set('')
        self.monthsTWE.set('')
        self.daysTWE.set('')


    def calc(self):
        self.clean()

        total = list()

        result = self.calcTotal(self.normalWE, 30)
        total.append(result)
        self.yearsTNWE.set(result[0])
        self.monthsTNWE.set(result[1])
        self.daysTNWE.set(result[2])

        result = self.calcTotal(self.hourlyPaidWE, 25)
        total.append(result)
        self.yearsTHPWE.set(result[0])
        self.monthsTHPWE.set(result[1])
        self.daysTHPWE.set(result[2])

        result = self.calcTotal(total, 30)
        self.yearsTWE.set(result[0])
        self.monthsTWE.set(result[1])
        self.daysTWE.set(result[2])


    def create_widgets(self):
        self.fWECalc = Frame(self.window)

        Label(self.fWECalc, text="\nΈτη:").grid(column=1, row=0, padx=10, pady=1)
        Label(self.fWECalc, text="\nΜήνες:").grid(column=2, row=0, padx=10, pady=1)
        Label(self.fWECalc, text="\nΗμέρες:").grid(column=3, row=0, padx=10, pady=1)

        Label(self.fWECalc, text="Νέα εγγραφή:").grid(column=0, row=1, padx=10, pady=10, sticky=E)
        Label(self.fWECalc, text="Καταχωρισμένες:").grid(column=0, row=3, padx=10, pady=1, sticky=E)
        Label(self.fWECalc, text="Κανονικές\nΑ/Α ΕΕ-ΜΜ-ΗΗ", justify=CENTER).grid(column=1, row=3, columnspan=2, padx=1, pady=1)
        Label(self.fWECalc, text="Ωρομίσθιες\nΑ/Α ΕΕ-ΜΜ-ΗΗ", justify=CENTER).grid(column=3, row=3, columnspan=2, padx=1, pady=1)

        Label(self.fWECalc, text="\nΈτη:").grid(column=1, row=6, padx=10, pady=1)
        Label(self.fWECalc, text="\nΜήνες:").grid(column=2, row=6, padx=10, pady=1)
        Label(self.fWECalc, text="\nΗμέρες:").grid(column=3, row=6, padx=10, pady=1)

        Label(self.fWECalc, text="Συνολική (κανονικές):").grid(column=0, row=7, padx=1, pady=10, sticky=E)
        Label(self.fWECalc, text="Συνολική (ωρομίσθιες):").grid(column=0, row=8, padx=1, pady=10, sticky=E)
        Label(self.fWECalc, text="Συνολική Προϋπηρεσία:").grid(column=0, row=9, padx=1, pady=10, sticky=E)

        self.year = StringVar()
        self.ntrYear = Entry(self.fWECalc, width=10, textvariable=self.year, justify=CENTER)
        self.ntrYear.grid(column=1, row=1, padx=10, pady=10)

        self.month = StringVar()
        self.ntrMonth = Entry(self.fWECalc, width=10, textvariable=self.month, justify=CENTER)
        self.ntrMonth.grid(column=2, row=1, padx=10, pady=10)

        self.day = StringVar()
        self.ntrDay = Entry(self.fWECalc, width=10, textvariable=self.day, justify=CENTER)
        self.ntrDay.grid(column=3, row=1, padx=10, pady=10)

        self.hourlyPaid = BooleanVar()
        self.ckbHourlyPaid = Checkbutton(self.fWECalc, text="Ωρομίσθιος", variable=self.hourlyPaid)
        self.ckbHourlyPaid.grid(column=4, row=1, padx=10, pady=10)

        self.btnAddWE = Button(self.fWECalc, text="Καταχώριση", command=self.addWE)
        self.btnAddWE.grid(column=0, row=2, columnspan=5, padx=10, pady=10)

        self.stNormalWE = ScrolledText(self.fWECalc, width=20, height=10, state='disabled')
        self.stNormalWE.grid(column=1, row=4, columnspan=2, padx=10, pady=10)

        self.stHourlyPaidWE = ScrolledText(self.fWECalc, width=20, height=10, state='disabled')
        self.stHourlyPaidWE.grid(column=3, row=4, columnspan=2, padx=10, pady=10)

        self.btnCalc = Button(self.fWECalc, text="Yπολογισμός", command=self.calc)
        self.btnCalc.grid(column=0, row=5, columnspan=5, padx=10, pady=10)

        self.yearsTNWE = StringVar()
        self.ntrYearsTNWE = Entry(self.fWECalc, width=10, textvariable=self.yearsTNWE, justify=CENTER, state='readonly')
        self.ntrYearsTNWE.grid(column=1, row=7, padx=10, pady=10)

        self.monthsTNWE = StringVar()
        self.ntrMonthsTNWE = Entry(self.fWECalc, width=10, textvariable=self.monthsTNWE, justify=CENTER, state='readonly')
        self.ntrMonthsTNWE.grid(column=2, row=7, padx=10, pady=10)

        self.daysTNWE = StringVar()
        self.ntrDaysTNWE = Entry(self.fWECalc, width=10, textvariable=self.daysTNWE, justify=CENTER, state='readonly')
        self.ntrDaysTNWE.grid(column=3, row=7, padx=10, pady=10)

        self.yearsTHPWE = StringVar()
        self.ntrYearsTHPWE = Entry(self.fWECalc, width=10, textvariable=self.yearsTHPWE, justify=CENTER, state='readonly')
        self.ntrYearsTHPWE.grid(column=1, row=8, padx=10, pady=10)

        self.monthsTHPWE = StringVar()
        self.ntrMonthsTHPWE = Entry(self.fWECalc, width=10, textvariable=self.monthsTHPWE, justify=CENTER, state='readonly')
        self.ntrMonthsTHPWE.grid(column=2, row=8, padx=10, pady=10)

        self.daysTHPWE = StringVar()
        self.ntrDaysTHPWE = Entry(self.fWECalc, width=10, textvariable=self.daysTHPWE, justify=CENTER, state='readonly')
        self.ntrDaysTHPWE.grid(column=3, row=8, padx=10, pady=10)

        self.yearsTWE = StringVar()
        self.ntrYearsTWE = Entry(self.fWECalc, width=10, textvariable=self.yearsTWE, justify=CENTER, state='readonly')
        self.ntrYearsTWE.grid(column=1, row=9, padx=10, pady=10)

        self.monthsTWE = StringVar()
        self.ntrMonthsTWE = Entry(self.fWECalc, width=10, textvariable=self.monthsTWE, justify=CENTER, state='readonly')
        self.ntrMonthsTWE.grid(column=2, row=9, padx=10, pady=10)

        self.daysTWE = StringVar()
        self.ntrDaysTWE = Entry(self.fWECalc, width=10, textvariable=self.daysTWE, justify=CENTER, state='readonly')
        self.ntrDaysTWE.grid(column=3, row=9, padx=10, pady=10)

        self.fWECalc.pack()


gui = GUI()
gui.window.mainloop()

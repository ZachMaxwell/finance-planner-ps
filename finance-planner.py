import openpyxl

path_to_sheet = r" "
working_sheet = input("Input the monthly sheet name:\n> ")
absolute_path = path_to_sheet + "\\" + working_sheet

wb = openpyxl.load_workbook(absolute_path)
sheet = wb.active

sheet['A1'] = 'Miscellaneous Expenses'
sheet['B1'] = 'Going Out/Eating Out Expenses'
sheet['C1'] = 'Grocery Expenses'
sheet['D1'] = 'Gas Expenses'
sheet['E1'] = 'Car/Auto Insurance Expenses'
sheet['F1'] = 'Subscriptions'
sheet['G1'] = 'Loan Payments'
sheet['H1'] = 'Rent/Utilities'
sheet['I1'] = 'Investment Contributions'
sheet['J1'] = 'Total Income'
sheet['K1'] = 'Total Expenses'
sheet['L1'] = 'Monthly Net Gain or Loss'

sheet['P1'] = 'Misc Totals'
sheet['P2'] = 'Go/Eo Totals'
sheet['P3'] = 'Grocery Totals'
sheet['P4'] = 'Gas Totals'
sheet['P5'] = 'Car Totals'
sheet['P6'] = 'Subscriptions Totals'
sheet['P7'] = 'Loans Totals'
sheet['P8'] = 'Rent/Utilities Totals'
sheet['P9'] = 'Investment Totals'

def add_expense():
    accepted_responses = ('y','n')
    contin = input("Add an expense? (y or n)\n> ")
    while contin not in accepted_responses:
        print("Try again.\n")
        contin = input("Add an expense? (y or n)\n> ")

    while contin == accepted_responses[0]:
        category = input('Pick a category\n a - Misc\n b - Go/Eo\n c - Grocery\n d - Gas\n e - Car\n f - Subscriptions\n g - Loans\n h - Rent & Utilities\n i - Investments\n> ')
        float_expense = input("Input an expense\n> ")

        if category == 'a':
            for count, row in enumerate(sheet['A'], 1):
                print(count, row)
            sheet['A' + str(count + 1)] = float(float_expense)

        elif category == 'b':
            for count, row in enumerate(sheet['B'], 1):
                print(count, row)
            sheet['B' + str(count + 1)] = float(float_expense)

        elif category == 'c':
            for count, row in enumerate(sheet['C'], 1):
                print(count, row)
            sheet['C' + str(count + 1)] = float(float_expense)

        elif category == 'd':
            for count, row in enumerate(sheet['D'], 1):
                print(count, row)
            sheet['D' + str(count + 1)] = float(float_expense)

        elif category == 'e':
            for count, row in enumerate(sheet['E'], 1):
                print(count, row)
            sheet['E' + str(count + 1)] = float(float_expense)

        elif category == 'f':
            for count, row in enumerate(sheet['F'], 1):
                print(count, row)
            sheet['F' + str(count + 1)] = float(float_expense)

        elif category == 'g':
            for count, row in enumerate(sheet['G'], 1):
                print(count, row)
            sheet['G' + str(count + 1)] = float(float_expense)

        elif category == 'h':
            for count, row in enumerate(sheet['H'], 1):
                print(count, row)
            sheet['H' + str(count + 1)] = float(float_expense)

        elif category == 'i':
            for count, row in enumerate(sheet['I'], 1):
                print(count, row)
            sheet['I' + str(count + 1)] = float(float_expense)

        contin = input("Add another expense? (y or n)\n> ")
        while contin not in accepted_responses:
            print("Try again.\n")
            contin = input("Add another expense? (y or n)\n> ")


    category_subtotals = {'A': 0, 'B': 0, 'C': 0, 'D': 0, 'E': 0, 'F': 0, 'G': 0, 'H': 0, 'I': 0}
    for row in sheet['A']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['A'] += float(row.value)

    for row in sheet['B']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['B'] += float(row.value)

    for row in sheet['C']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['C'] += float(row.value)

    for row in sheet['D']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['D'] += float(row.value)

    for row in sheet['E']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['E'] += float(row.value)

    for row in sheet['F']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['F'] += float(row.value)

    for row in sheet['G']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['G'] += float(row.value)

    for row in sheet['H']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['H'] += float(row.value)

    for row in sheet['I']:
        if type(row.value) == float or type(row.value) == int:
            category_subtotals['I'] += float(row.value)

    sheet['Q1'] = category_subtotals['A']
    sheet['Q2'] = category_subtotals['B']
    sheet['Q3'] = category_subtotals['C']
    sheet['Q4'] = category_subtotals['D']
    sheet['Q5'] = category_subtotals['E']
    sheet['Q6'] = category_subtotals['F']
    sheet['Q7'] = category_subtotals['G']
    sheet['Q8'] = category_subtotals['H']
    sheet['Q9'] = category_subtotals['I']

    return category_subtotals

def add_deposit():
    accepted_responses = ('y','n')
    contin = input("Add a deposit? (y or n)\n> ")
    while contin not in accepted_responses:
        print("Try again.\n")
        contin = input("Add a deposit? (y or n)\n> ")

    while contin == accepted_responses[0]:
        float_income = input("Input a deposit\n> ")

        for count, row in enumerate(sheet['J'], 1):
            print(count, row)
        sheet['J' + str(count + 1)] = float(float_income)

        contin = input("Add another deposit?\n> ")
        while contin not in accepted_responses:
            print("Try again.\n")
            contin = input("Add another deposit? (y or n)\n> ")

def calculate_totals(category_subtot_dictionary):
    running_expense_total = 0
    for val in category_subtot_dictionary.values():
        running_expense_total += val
    sheet['K2'] = float(running_expense_total)

    income_expense_subtotal = {'J': 0, 'K': 0}
    for row in sheet['J']:
        if type(row.value) == float or type(row.value) == int:
            income_expense_subtotal['J'] += float(row.value)

    for row in sheet['K']:
        if type(row.value) == float or type(row.value) == int:
            income_expense_subtotal['K'] += float(row.value)

    sheet['L2'] = income_expense_subtotal['J'] - income_expense_subtotal['K']


expense_output = add_expense()
add_deposit()
calculate_totals(expense_output)

wb.save(absolute_path)

print("*** Sheet has been updated and saved ***")

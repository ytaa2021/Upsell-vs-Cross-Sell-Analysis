import csv
import openpyxl


file = r"C:\Users\yotam.twersky\Downloads\cross_sell_vs_upsell (2).csv.xlsx"

wb = openpyxl.load_workbook(file)
ws = wb.active
cell_range = 'A2:O19233'
data = [[cell.value for cell in row] for row in ws[cell_range]]


data.sort(key=lambda x: x[1])
for i in data:
    
    try:
        #filter data for all items that are hte same company
        scmp = list(filter(lambda x: x[1] == i[1], data))
        #filter those items for previous transactions
        prts = list(filter(lambda x: x[6] < i[6], scmp))
    except:
        prts = []
    #filter those items for items with the same product codes
    if prts == []:
        i.append(["no previous transactions found"])
    else:
        proco = list(filter(lambda x: x[4] == i[4], prts))
        if proco != []:
            proco.sort(key=lambda x: x[6])
            if proco[-1][10] < i[10] and proco[-1][14] < i[14]:
                i.append(["most likely upsell"])
                #if the current number of seats is greater than the past & the amount of money is greater than the past
                #it is most likely upsell
            elif proco[-1][14] < i[14]:
                i.append(["potentially upsell"])
                #the amount of money is greater than the past
                #it is possibly upsell
            else:
                i.append(["same product code but likely not upsell"])
        else:
            i.append(["likely cross sell"])

doofile = open(r"C:\Users\yotam.twersky\OneDrive - Kiteworks\Desktop\exported_cross_upsell.csv", 'w')
csv_writer = csv.writer(doofile)
csv_writer.writerows(data)
            



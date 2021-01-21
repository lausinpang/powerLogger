import os
from os.path import isfile, join
import csv
from matplotlib import pyplot as plt
from itertools import islice
import matplotlib.ticker as plticker
import numpy as np



def get_row_nums(opt):
    switcher = {
        "A": ( "Current [A]", (13, 16, 19), ("Current rms L1", "Current rms L2", "Current rms L3"), 240.00 ), # choice : ( y-label, column numbers in excel, line legends, max )
        "B": ( "Voltage [V]", (4, 7, 10), ("Voltage rms L1", "Voltage rms L2", "Voltage rms L3"), 280.00 ),
        "C": ( "Power factor", (22, 23, 24) ),
        "D": ( "THD-F [%]", () , ("THD-F Current L1", "THD-F Current L2", "THD-F Current L3"), 30.00)
    }

    return switcher.get(opt, "Invalid input")


def save_graph(filename, opt):
    opt_row = get_row_nums(opt)
    x = []
    y1 = []
    y2 = []
    y3 = []
    with open(filename,'r') as csvfile:
        # plots = csv.reader(csvfile, delimiter=',')
        for row in islice(csv.reader(csvfile), 18, None):
            try:
                dateTime = row[0].split(" ")
                eTime = row[1]
                if(row[1]=="STOP"):
                    break
                x_value = dateTime[0] + "\n" + dateTime[1]
                x.append(x_value)
                y1.append(float( row[opt_row[1][0]])) #I1_Avg[A]
                y2.append(float( row[opt_row[1][1]])) #I2_Avg[A]
                y3.append(float( row[opt_row[1][2]])) #I3_Avg[A]
            except:
                print(f'Check row {eTime}')

    # graph properties
    cm = 1/2.54  # centimeters in inches
    fig, ax = plt.subplots(figsize=(15*3*cm, 4.5*3*cm))

    # plot X & Y
    ax.plot(x, y1, label = opt_row[2][0], color="black")
    ax.plot(x, y2, label = opt_row[2][1], color="red")
    ax.plot(x, y3, label = opt_row[2][2], color="blue")

    # y = np.random.randint(low=0, high=opt_row[3])
    plt.yticks(np.arange(0, opt_row[3], 40))

    # set modify X-axis display and intervals
    loc = plticker.MultipleLocator(base=28)
    ax.xaxis.set_major_locator(loc)
    plt.xticks(fontsize=6)

    # finalize graph
    plt.ylabel(opt_row[0])
    plt.grid()
    plt.gca().legend(loc='lower center', bbox_to_anchor=(0.5, -0.25), ncol=3, columnspacing=15)
    plt.gcf().autofmt_xdate(rotation=0,ha='center')
    # plt.show()

    plt.savefig('graphs/'+opt_row[0]+"_"+os.path.basename(filename)+'.png')


def homescreen():
    print("\n\nChoose information to plot: "+
            "\n\n[A] Current"+
            "\n[B] Voltage"+
            "\n[C] Power Factor (not yet available)"+
            "\n[D] Total Harmonic Distribution (not yet available)")
    choice = input("Enter your choice: ").upper()

    return choice



choice = homescreen()
path = 'files/'
for file in os.listdir(path):
    if isfile(join(path, file)):
        save_graph(join(path, file), choice)
print("DONE")
# save_graph('file.csv')


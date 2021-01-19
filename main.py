import os
from os.path import isfile, join
import csv
from matplotlib import pyplot as plt
from itertools import islice
import matplotlib.ticker as plticker



def save_graph(filename):
    x1 = []
    y1 = []
    x2 = []
    y2 = []
    x3 = []
    y3 = []
    with open('file.csv','r') as csvfile:
        # plots = csv.reader(csvfile, delimiter=',')
        for row in islice(csv.reader(csvfile), 18, None):
            try:
                dateTime = row[0].split(" ")
                eTime = row[1]
                if(row[1]=="STOP"):
                    break
                x_value = dateTime[0] + "\n" + dateTime[1]
                x1.append(x_value)
                y1.append(float( row[13])) #I1_Avg[A]
                x2.append(x_value)
                y2.append(float( row[16])) #I2_Avg[A]
                x3.append(x_value)
                y3.append(float( row[19])) #I3_Avg[A]
            except:
                print(f'Check row {eTime}')

    # graph properties
    cm = 1/2.54  # centimeters in inches
    fig, ax = plt.subplots(figsize=(15*3*cm, 4.5*3*cm))

    # plot X & Y
    ax.plot(x1, y1, label = "Current rms L1", color="black")
    ax.plot(x2, y2, label = "Current rms L2", color="red")
    ax.plot(x3, y3, label = "Current rms L3", color="blue")

    # set modify X-axis display and intervals
    loc = plticker.MultipleLocator(base=28)
    ax.xaxis.set_major_locator(loc)
    plt.xticks(fontsize=6)

    # finalize graph
    plt.ylabel('Current [A]')
    plt.gca().legend(loc='lower center', bbox_to_anchor=(0.5, -0.25), ncol=3, columnspacing=15)
    plt.gcf().autofmt_xdate(rotation=0,ha='center')
    # plt.show()

    plt.savefig('graphs/'+os.path.splitext(filename)[0]+'.png')



path = 'files/'
for file in os.listdir(path):
    if isfile(join(path, file)):
        # print(os.path.basename(file))
        save_graph(file)
print("DONE")
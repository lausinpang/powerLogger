import os
from os.path import isfile, join
from matplotlib import pyplot as plt
from docx import Document
from docx.shared import RGBColor
from docx.shared import Inches
import matplotlib.ticker as plticker
import pandas as pd
import numpy as np
from datetime import datetime, timedelta


cm = 1/2.54  # centimeters in inches
exclude_rows = [1, 2, 3, 4, 5, 6]  # rows to exclude in reading CSV file
graph_types = ["Current [A]", "THD-F [%]", "Power factor", "Voltage [V]"]  # y axis
# periods = ["weekdays", "all", "saturdays"] #measurement period/x axis in graph


def sitename():
    sitename = input("Enter the site shortform: ")

    return sitename


def save_to_png(plt, file, opt):
	script_dir = os.path.dirname(__file__)
	results_dir = os.path.join(script_dir, 'graphs/')
	if not os.path.isdir(results_dir):
		os.makedirs(results_dir)
	plt.savefig('graphs/'+opt+"_"+os.path.splitext(file)[0]+'.png')


def graph_current_weekdays():
	pass
# code for current in weekdays period


def graph_current_sat():
	pass
# code for current in saturdays


def graph_meas_period(columns, opt, file):
	# plot graph
	new_df = df[["dateTime"] + columns]
	new_df.plot(x="dateTime", y=columns, figsize=(15*3*cm, 5*3*cm))

	if opt != "Power factor":
		legend = plt.legend(bbox_to_anchor=(0.8, -0.05), ncol=3, columnspacing=15)

		# set legend text
		for col in columns:
			legend.get_texts()[columns.index(col)].set_text(opt.split(" ")[0]+ " " + col)

	# setup graph
	plt.xticks(fontsize=6)
	plt.grid()
	plt.xlabel("")
	plt.ylabel(opt)

	save_to_png(plt, file, opt)

	idx = new_df.index
	num_of_rows = len(idx)
	return num_of_rows


def graph_cur_vs_thd(cols_current, cols_thd):
	graph_df = df[cols_current+cols_thd]
	ax1= graph_df.plot(kind='scatter', x="I1[A]", y="THD-F_I1[%]", figsize=(15*3*cm, 5*3*cm),
					label="Phase 1", color='r')
	ax2= graph_df.plot(kind='scatter', x="I2[A]", y="THD-F_I2[%]", figsize=(15*3*cm, 5*3*cm),
					label="Phase 2", ax=ax1, color='g')
	ax3= graph_df.plot(kind='scatter', x="I3[A]", y="THD-F_I3[%]", figsize=(15*3*cm, 5*3*cm),
					label="Phase 3", ax=ax1, color='b')

	# setup graph
	plt.xticks(fontsize=6)
	plt.grid()
	plt.xlabel("Phase Current (A)")
	plt.ylabel("Current Harmonics (%)")

	save_to_png(plt, file, "Current_vs_THD")


def get_columns(opt):
	switcher = {
		"Current [A]": 	["I1[A]", "I2[A]", "I3[A]"],
		"THD-F [%]": 	["THD-F_I1[%]", "THD-F_I2[%]", "THD-F_I3[%]"],
		"Power factor": ["PF"],
		"Voltage [V]": 	["U1[V]", "U2[V]", "U3[V]"]
	}
	return switcher.get(opt, "Invalid input")


def get_df(file):  # get dataFrame from csv
	data = pd.read_csv(file, skiprows=lambda x: x in exclude_rows, header=0)
	new_data = data.assign(dateTime=data.Date + "\n" + data.Time)
	return new_data


def get_max(df, col):
	max_index = df[col].idxmax()
	return df[col].max(), df.iloc[max_index, 0] + ' ' + df.iloc[max_index, 1]


def get_min(df, col):
	min_index = df[col].idxmin()
	return df[col].min(), df.iloc[min_index, 0] + ' ' + df.iloc[min_index, 1]

def get_average(df):
	average_THD1 = round(df['THD-F_I1[%]'].mean(skipna=True), 2)
	average_THD2 = round(df['THD-F_I2[%]'].mean(skipna=True), 2)
	average_THD3 = round(df['THD-F_I3[%]'].mean(skipna=True), 2)
	average_i1 = round(df['I1[A]'].mean(skipna=True), 2)
	average_i2 = round(df['I2[A]'].mean(skipna=True), 2)
	average_i3 = round(df['I3[A]'].mean(skipna=True), 2)
	average_PF = round(df['PF'].mean(skipna=True), 2)
	average_v1 = round(df['U1[V]'].mean(skipna=True), 2)
	average_v2 = round(df['U2[V]'].mean(skipna=True), 2)
	average_v3 = round(df['U3[V]'].mean(skipna=True), 2)
	details = {
		"AVE i1": average_i1,
		"AVE i2": average_i2,
		"AVE i3": average_i3,
		"AVE v1": average_v1,
		"AVE v2": average_v2,
		"AVE v3": average_v3,
		"AVE PF": average_PF,
		"AVE THD1": average_THD1,
		"AVE THD2": average_THD2,
		"AVE THD3": average_THD3,
	}
	return details


def get_circuit_details(data, file):
	details = {
			"Circuit Name": os.path.splitext(file)[0].split("_")[1],
			"Circuit current": os.path.splitext(file)[0].split("_")[2],
			"THD 1 MAX": get_max(data, "THD-F_I1[%]"),
			"THD 1 MIN": get_min(data, "THD-F_I1[%]"),
			"THD 2 MAX": get_max(data, "THD-F_I2[%]"),
			"THD 2 MIN": get_min(data, "THD-F_I2[%]"),
			"THD 3 MAX": get_max(data, "THD-F_I3[%]"),
			"THD 3 MIN": get_min(data, "THD-F_I3[%]"),

			"CUR 1 MAX": get_max(data, "I1[A]"),
			"CUR 1 MIN": get_min(data, "I1[A]"),
			"CUR 2 MAX": get_max(data, "I2[A]"),
			"CUR 2 MIN": get_min(data, "I2[A]"),
			"CUR 3 MAX": get_max(data, "I3[A]"),
			"CUR 3 MIN": get_min(data, "I3[A]"),

			"VOL 1 MAX": get_max(data, "U1[V]"),
			"VOL 1 MIN": get_min(data, "U1[V]"),
			"VOL 2 MAX": get_max(data, "U2[V]"),
			"VOL 2 MIN": get_min(data, "U2[V]"),
			"VOL 3 MAX": get_max(data, "U3[V]"),
			"VOL 3 MIN": get_min(data, "U3[V]"),

			"PF MAX": get_max(data, "PF"),
			"PF MIN": get_min(data, "PF"),
	}
	return details


number = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
		   "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]


# MAIN #
circuit_total = 0  # total number of circuit (excel)
incomer_total = 0  # total number of incomer (incomer will container in the excel name)
current_details = {}  # each current name and their rated current (e.g. kitty incomer_3200)
number_rows = []
site_name = sitename()

circuit_details = list()
circuit_average = list()
# max and min and their coorsponding time or current, voltage, power factor and THD
# average of current, voltage, power factor and THD

path = 'files/'
# LOOP for each circuit/excel file
for file in os.listdir(path):
	if isfile(join(path, file)):
		circuit_total += 1
		count = 0
		test = "Incomer"
		count = file.count(test)
		if count == 1:
			incomer_total += 1
		df = get_df(join(path, file))
		circuit_details.append(get_circuit_details(df, file))
		circuit_average.append(get_average(df))

		for option in graph_types:  # loop to graph each type
			col_list = get_columns(option)
			num_rows = graph_meas_period(col_list, option, file)
			number_rows.append(str(num_rows))
			#if (option == "Current [A]"):
				# graph_current_weekday()
				# graph_current_sat()

		graph_cur_vs_thd(get_columns("Current [A]"), get_columns("THD-F [%]"))

# write circuit info to file
# circuits_df = pd.DataFrame.from_dict(circuit_details)
# circuits_df.to_excel('circuit_details.xlsx')

print(f'Number of circuits: {circuit_total}\nTotal Incomer: {incomer_total}\n')
range(len(circuit_details))
range(len(circuit_average))
numbers = str(incomer_total)
circuit400 = circuit_total - incomer_total
circuit_number = str(circuit400)
# print(circuit_details[0]["Circuit Name"])

document = Document()
document.add_heading('Power Quality Analysis', 0)
p = document.add_paragraph(site_name)
p.add_run(' is served by ')
p.add_run(number[incomer_total])
p.add_run(circuit_details[0]["Circuit current"])
p.add_run('A incomers. ')
p.add_run('Energenz conducted power quality measurements each of the ')
p.add_run(number[incomer_total])
p.add_run(' of the incomers and')
p.add_run(number[circuit400])
p.add_run(' (' + circuit_number + ') ')
p.add_run('400A subcircuits:')

for x in range(circuit_total):
	p = document.add_paragraph(circuit_details[x]["Circuit current"], style='List Bullet')
	p.add_run('A - ')
	p.add_run(circuit_details[x]["Circuit Name"])

document.add_paragraph('The section below presented the voltage (V), current (A), power factor (PF), ' +
					   'and total harmonic distortion (THD) findings of each circuit.')

for x in range(incomer_total):
	document.add_heading(circuit_details[x]["Circuit Name"] + ' (' + circuit_details[x]["Circuit current"] + 'A)', level=1)
	p = document.add_paragraph().add_run('Current and Voltage')
	p.bold = True
	a = document.add_paragraph('The')
	a.add_run('(' + ')')
	a.add_run('serves')

	# add current table and label format
	table = document.add_table(rows=4, cols=5)
	table.style = 'Colorful List Accent 6'
	label_cell0 = table.cell(0, 0)
	label0_paragraph = label_cell0.paragraphs[0]
	run0 = label0_paragraph.add_run('Parameter')
	run0.font.color.rgb = RGBColor(255, 255, 255)
	label_cell1 = table.cell(0, 1)
	label1_paragraph = label_cell1.paragraphs[0]
	run1 = label1_paragraph.add_run('Times')
	run1.font.color.rgb = RGBColor(255, 255, 255)
	label_cell2 = table.cell(0, 2)
	label2_paragraph = label_cell2.paragraphs[0]
	run2 = label2_paragraph.add_run('Current RMS\n Max.')
	run2.font.color.rgb = RGBColor(255, 255, 255)
	label_cell3 = table.cell(0, 3)
	label3_paragraph = label_cell3.paragraphs[0]
	run3 = label3_paragraph.add_run('Current RMS\n Min.')
	run3.font.color.rgb = RGBColor(255, 255, 255)
	label_cell4 = table.cell(0, 4)
	label4_paragraph = label_cell4.paragraphs[0]
	run4 = label4_paragraph.add_run('Current RMS Average')
	run4.font.color.rgb = RGBColor(255, 255, 255)

	# information of the circuit
	first_cells = table.rows[1].cells
	first_cells[0].text = circuit_details[x]["Circuit Name"] + ' I1 [A]'
	first_cells[1].text = number_rows[x]
	curmax = str(circuit_details[x]["CUR 1 MAX"])
	first_cells[2].text = curmax
	curmin = str(circuit_details[x]["CUR 1 MIN"])
	first_cells[3].text = curmin
	curave = str(circuit_average[x]["AVE i1"])
	first_cells[4].text = curave
	second_cells = table.rows[2].cells
	second_cells[0].text = circuit_details[x]["Circuit Name"] + ' I2 [A]'
	second_cells[1].text = number_rows[x]
	curmax = str(circuit_details[x]["CUR 2 MAX"])
	second_cells[2].text = curmax
	curmin = str(circuit_details[x]["CUR 2 MIN"])
	second_cells[3].text = curmin
	curave = str(circuit_average[x]["AVE i2"])
	second_cells[4].text = curave
	third_cells = table.rows[3].cells
	third_cells[0].text = circuit_details[x]["Circuit Name"] + ' I2 [A]'
	third_cells[1].text = number_rows[x]
	curmax = str(circuit_details[x]["CUR 3 MAX"])
	third_cells[2].text = curmax
	curmax = str(circuit_details[x]["CUR 3 MIN"])
	third_cells[3].text = curmin
	curave = str(circuit_average[x]["AVE i3"])
	third_cells[4].text = curave

	# add current of first circuit
	i = x+1
	filepath = os.path.abspath('graphs/Current [A]_' + str(i) + '_' + circuit_details[x]["Circuit Name"] + '_' + circuit_details[x]["Circuit current"] + '.png')
	document.add_picture(filepath, width=Inches(7.0))

	# add voltage table and label format
	table1 = document.add_table(rows=4, cols=5)
	table1.style = 'Colorful List Accent 6'
	label_cell0 = table1.cell(0, 0)
	label0_paragraph = label_cell0.paragraphs[0]
	run0 = label0_paragraph.add_run('Parameter')
	run0.font.color.rgb = RGBColor(255, 255, 255)
	label_cell1 = table1.cell(0, 1)
	label1_paragraph = label_cell1.paragraphs[0]
	run1 = label1_paragraph.add_run('Times')
	run1.font.color.rgb = RGBColor(255, 255, 255)
	label_cell2 = table1.cell(0, 2)
	label2_paragraph = label_cell2.paragraphs[0]
	run2 = label2_paragraph.add_run('Voltage RMS\n Max.')
	run2.font.color.rgb = RGBColor(255, 255, 255)
	label_cell3 = table1.cell(0, 3)
	label3_paragraph = label_cell3.paragraphs[0]
	run3 = label3_paragraph.add_run('Voltage RMS\n Min.')
	run3.font.color.rgb = RGBColor(255, 255, 255)
	label_cell4 = table1.cell(0, 4)
	label4_paragraph = label_cell4.paragraphs[0]
	run4 = label4_paragraph.add_run('Voltage RMS Average')
	run4.font.color.rgb = RGBColor(255, 255, 255)

	# information of the circuit
	first_cells = table1.rows[1].cells
	first_cells[0].text = circuit_details[x]["Circuit Name"] + ' U1 [V]'
	first_cells[1].text = number_rows[x]
	curmax = str(circuit_details[x]["VOL 1 MAX"])
	first_cells[2].text = curmax
	curmin = str(circuit_details[x]["VOL 1 MIN"])
	first_cells[3].text = curmin
	curave = str(circuit_average[x]["AVE v2"])
	first_cells[4].text = curave
	second_cells = table1.rows[2].cells
	second_cells[0].text = circuit_details[x]["Circuit Name"] + ' U2 [V]'
	second_cells[1].text = number_rows[x]
	curmax = str(circuit_details[x]["VOL 2 MAX"])
	second_cells[2].text = curmax
	curmin = str(circuit_details[x]["VOL 2 MIN"])
	second_cells[3].text = curmin
	curave = str(circuit_average[x]["AVE v2"])
	second_cells[4].text = curave
	third_cells = table1.rows[3].cells
	third_cells[0].text = circuit_details[x]["Circuit Name"] + ' U2 [V]'
	third_cells[1].text = 'times'
	curmax = str(circuit_details[x]["VOL 3 MAX"])
	third_cells[2].text = curmax
	curmin = str(circuit_details[x]["VOL 3 MIN"])
	third_cells[3].text = curmin
	curave = str(circuit_average[x]["AVE v3"])
	third_cells[4].text = curave

	# add current of first voltage
	i = x + 1
	filepath = os.path.abspath(
		'graphs/Voltage [V]_' + str(i) + '_' + circuit_details[x]["Circuit Name"] + '_' + circuit_details[x][
			"Circuit current"] + '.png')
	document.add_picture(filepath, width=Inches(7.0))

	document.add_paragraph()
	p = document.add_paragraph().add_run('Total Harmonic Distortion')
	p.bold = True
	document.add_paragraph('The Code of Practice (COP) of Energy Efficiency of Building Services Installation has ' +
						   'regulated the maximum THD at difference rated circuit current. The details are shown below.')
	if int(circuit_details[x]["Circuit current"])>= 400 and int(circuit_details[x]["Circuit current"]) < 800:
		circuitTHD = "12.0%"
		wording = "is between 400A and 800A"
	elif int(circuit_details[x]["Circuit current"]) >= 800 and int(circuit_details[x]["Circuit current"]) < 2000:
		circuitTHD = "8.0%"
		wording = "is between 800A and 2,000A"
	elif int(circuit_details[x]["Circuit current"]) > 2000:
		circuitTHD = "5.0%"
		wording = "is greater than 2,000A"

	first = document.add_paragraph('As shown in the table, the rated current of the incomer ' + wording)
	first.add_run(' and the harmonic distortion percentage of current should not be above ' + circuitTHD + '. ')
	first.add_run('The max. current THD is around')
	first.add_run('Overall, the current harmonic distortion percentage of the measurement period ')
	first.add_run(
		'satisfy the requirement of the Code of Practice of Energy Efficiency of Building Services Installation.')

	p = document.add_paragraph().add_run('Figure presents the THD current result for')

	i = x + 1
	filepath = os.path.abspath(
		'graphs/THD-F [%]_' + str(i) + '_' + circuit_details[x]["Circuit Name"] + '_' + circuit_details[x][
			"Circuit current"] + '.png')
	document.add_picture(filepath, width=Inches(7.0))


document.save('Power Quality Analysis.docx')

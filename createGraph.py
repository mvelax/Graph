import xlsxwriter
import csv
import glob
from xlsxwriter.utility import xl_col_to_name

#Read a csv file and create graphs

STARTING_GRAPH_COLUMN = 1

def addDefaultColumnSeries(chart, sheet_name, y2_axis, column, last_row):
	STARTING_ROW = 1
	HEADER = 0
	"""
	Add a series using the header of that column,
	values will use row 2 to the one determined in rows.
	leftmost column will determine the category.
	chart: chart object where series will be added.
	sheet_name: string containing the sheet name.
	y2_axis: boolean, states if series should be part of y_axis 1(False) or 2(True).
	column: 0 index numer indicating the columnd witht the desired data.
	last_row: number indicating how many 
	"""
	if not y2_axis:
		chart.add_series({
			"name":			[sheet_name, HEADER,column],
			"categories":	[sheet_name, STARTING_ROW, HEADER, last_row, HEADER],
			"values":		[sheet_name, STARTING_ROW, column, last_row,column]
		})
	else:
		chart.add_series({
			"name":			[sheet_name, HEADER,column],
			"categories":	[sheet_name, STARTING_ROW, HEADER, last_row, HEADER],
			"values":		[sheet_name, STARTING_ROW, column, last_row,column],
			"y2_axis": 1
		})
		
	
	
def makeChart(workbook, data_worksheet, info, last_row):
	"""
	Creates a chart in the worksheet, inspired in grafico1.
	info is a tuple like:
	([y1_axis_cols],[y2_axis_cols],number,name)
	"""
	chart_sheet = workbook.add_chartsheet("Grafico{}".format(info[2]))
	chart = workbook.add_chart({"type":"line"})
	addANRExecChart(chart)
	for column in info[0]:
		addDefaultColumnSeries(chart, data_worksheet.get_name(), False, column, last_row)
	for column in info[1]:
		addDefaultColumnSeries(chart, data_worksheet.get_name(), True, column, last_row)
	chart.set_legend({'position': 'bottom'})
	chart.set_title({"name":info[3]})
	chart_sheet.set_chart(chart)
	
	
def addANRExecChart(chart):
	pass

#Use user input
input_file = "test input.csv"
INPUT = "*.csv"
csv_file_list = glob.glob(INPUT)
ERICSSON_GRAPHS = [
			([2,4,14],[],1,"Volumen de trafico de voz cursado & Tasa de caidas de voz & Tasa de fallos de accesibilidad de voz"),
			([5],[6,12],2,"Volumen de trafico de datos & Tasa de fallos de accesibilidad & Tasa de caidas de datos"),
			([12,10],[],3,"Tasa de Accesibilidad HSDPA & Tasa Accesibilidad HSUPA"),
			([8],[],4,"Tasa de llamadas de voz originadas en 3G y que terminan en 2G"),
			([15],[16],5,"Volumen de SHO y Tasa de exito de SHO"),
			([17],[18],6,"Volumen de IFHO y Tasa de Fallos de IFHO")
		]
OFFSET = 2 #Difference between columns in Metricas_Datos and Helper Table.
			


for file_name in csv_file_list:
	file_name_split = file_name.split(".")
	assert(len(file_name_split) == 2)
	file_name_no_ext = file_name_split[0]

	wb = xlsxwriter.Workbook(file_name_no_ext+".xlsx", {'strings_to_numbers':  True})
	ws = wb.add_worksheet("Metricas_Datos") #Holds csv data
	helper_sheet = wb.add_worksheet("Helper Table") #Calculates pre and post ANR averages.

	with open(file_name, "r") as csv_file:
		csv_reader = csv.reader(csv_file)
		number_of_rows = 0
		header = []
		time_col = []
		for i,row in enumerate(csv_reader):
			time_col.append(row[0])
			if i==0:
				header = row
			clean_row = []
			for element in row:
				if element=="None":
					clean_row.append("")
				else:
					clean_row.append(element)
			ws.write_row(i,0,clean_row)
		else:
			number_of_rows = i+1
			number_of_columns = len(row)
			helper_sheet.write_column(0,0,time_col)
		
		for info in ERICSSON_GRAPHS:
			makeChart(wb,ws,info,number_of_rows)
		
		header = [header[0]] + ["ANR Execution", "ANR Execution", "ANR avg"] + header[2:]
		helper_sheet.write_row(0,0,header)
		avg_if_formula = "=AVERAGEIF($D:$D,$D{0},Metricas_Datos!{1}:{1})"
		for row in range(1,number_of_rows):
			for col in range(2,number_of_columns):
				col_letter = xl_col_to_name(col)
				helper_sheet.write(row,col+OFFSET,avg_if_formula.format(row+1,col_letter))

	wb.close()
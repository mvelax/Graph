import xlsxwriter
import csv
import glob

#Read a csv file and create graphs

STARTING_GRAPH_COLUMN = 1

def addDefaultColumnSeries(chart, sheet, column, last_row):
	STARTING_ROW = 1
	HEADER = 0
	"""
	Add a series using the header of that column,
	values will use row 2 to the one determined in rows.
	leftmost column will determine the category.
	"""
	chart.add_series({
		"name":			[sheet, HEADER,column],
		"categories":	[sheet, STARTING_ROW, HEADER, last_row, HEADER],
		"values":		[sheet, STARTING_ROW, column, last_row,column]
	})
	

#Use user input
input_file = "test input.csv"
INPUT = "*.csv"
csv_file_list = glob.glob(INPUT)

for file_name in csv_file_list:
	file_name_split = file_name.split(".")
	assert(len(file_name_split) == 2)
	file_name_no_ext = file_name_split[0]

	wb = xlsxwriter.Workbook(file_name_no_ext+".xlsx", {'strings_to_numbers':  True})
	ws = wb.add_worksheet("Metricas_Datos")

	chart = wb.add_chart({'type': 'line'})

	with open(file_name, "r") as csv_file:
		csv_reader = csv.reader(csv_file)
		number_of_rows = 0
		for i,row in enumerate(csv_reader):
			#Could store the values here into some lists, to process later.
			ws.write_row(i,0,row)
		else:
			number_of_rows = i+1
			number_of_columns = len(row)
		
		for column in range(STARTING_GRAPH_COLUMN,number_of_columns):
			addDefaultColumnSeries(chart, ws.get_name(), column, number_of_rows)

		ws.insert_chart(7,7, chart)

	wb.close()
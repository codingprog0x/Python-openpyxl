'''
######
#	Module for creating and saving an Excel file
######
'''

from openpyxl import Workbook
from openpyxl.formatting import Rule
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule

class CreateExcel:
	def __init__(self):
		######
		#	Create Workbook object, as well as worksheet
		######
		self.wb = Workbook()
		self.ws = self.wb.active
		self.text_attribute = Alignment(horizontal="center", vertical="bottom",
								text_rotation=0, wrap_text=False,
								shrink_to_fit=False, indent=0
							)
		self.bg_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
		
		######
		#	round_number tracks current Excel row
		######		
		self.round_number = 1
		
	######
	#	Change header dimensions
	######
	def update_cell_dimensions(self, height, *width):
		######
		#	Update in future to be more dynamic
		#	For now, usage case is very specific
		######
		self.ws.column_dimensions["A"].width = 15
		self.ws.column_dimensions["B"].width = 28
		self.ws.column_dimensions["C"].width = 28
		self.ws.column_dimensions["D"].width = 18
		self.ws.column_dimensions["E"].width = 18
		self.ws.column_dimensions["F"].width = 33
		
		self.ws.row_dimensions[1].height = height
	
	######
	#	Input cell values and format
	######	
	def update_cell(self, is_header=None, *contents):
		if is_header is None:
			print("Please specify True or False as to whether this update is a header.")
		cell_content = contents
		cells = []
		
		######
		#	Create and append a cell object to cells list
		#	Number of cell objects appended contingent on
		#	number of arguments provided
		#	Afterwards, increment round_number to move to next row
		######
		for x in range(len(cell_content)):
			cells.append(self.ws.cell(row=self.round_number, column=x+1, value=cell_content[x]))
			cells[x].alignment = self.text_attribute
			if is_header:
				cells[x].fill = self.bg_fill
		self.round_number += 1			

	def save_file(self, file_path=None, file_name=None):
		if file_path is None:
			print("\nPlease specify the file path.")
			return False
		
		if file_name is None:
			print("\nPlease specify the file name.")
			return False
	
		######
		#	Save file
		######
		try:
			self.wb.save(file_path + "/" + file_name)
		except IOError:
			print("Trying without '/' after file path.")
			try:
				self.wb.save(file_path + file_name)
			except IOError:
				print("\nX----X File saved unsuccessfully.")
				return False
		print("\n!----! File saved successfully.")

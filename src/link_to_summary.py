import xlwings as xw
from xlwings.constants import AutoFillType
import win32com.client as winclient
import helpers

class SummaryProcessor:
	def __init__(self,workbook) -> None:
		self.book = workbook
		self.summary_sht = workbook.sheets("Summary")
		self.all_sheets = workbook.sheets
		self.tabs = [s.name for s in self.all_sheets]
		
	main_keys = [
			'Allocation of Ordinary Income (Loss) per Partners',
			'Guaranteed Pymt',  #Placeholder, not mapped
			'Allocation of Rental Income (Loss) per Partners',
			'Allocation of Capital Gain (Loss) per Partners',
			]

	#Get list of tabs to be processed
	exclude_tabs = ["Summary","Summary Sample"]
	boundary_dict = {}

	#Connect to active excel App object with win32. 
	# Used to activate/deactive app properties e.g. screenupdating
	excelapp = winclient.gencache.EnsureDispatch('Excel.Application')

	WshShell = winclient.DispatchEx("WScript.Shell")

	def make_summaries(self):
		main_dict = {}
		names,sorted_tabs = self.get_unique_names()
		
		if sorted_tabs == None:
			self.toggle_app_state(val=True)
			return

		if len(names) == 0:
			self.WshShell.Popup("No names found on column A in any of the tabs", 10, "Missing info", 0+16)  #0=OK button, 16=critical
			self.toggle_app_state(val=True)
			return

		sorted_tabs,sorted_states = helpers.sort_tabs(sorted_tabs)
		for sheet_name in sorted_tabs:
			sht = self.book.sheets[sheet_name]
			boundary_rows = self.boundary_dict[sheet_name]
			
			main_dict[sheet_name] = {}
			main_dict[sheet_name]['main'] = self.create_main_dict(sht,boundary_rows) 

		#Create final list
		data_list,final_headings,final_states,boundary_columns = helpers.create_final_list(
						names=names, 
						dictionary = main_dict, 
						tabs = sorted_tabs,
						main_keys=self.main_keys,
						states = sorted_states
						)	

		#Fill in data and add styling
		self.finalise(
				tab_names=final_headings, 
				state_names = final_states, 
				names=names, 
				data_list=data_list,
				boundary_columns = boundary_columns
			)

		self.toggle_app_state(val=True)  #turn on screen updating


	def get_section_boundaries(self,sheet) -> list:
		"""
		Identifies the rows where each 'main_keys' section starts
		"""
		col_a = sheet.range("A1:A300").value
		boundary_rows = []
		for key in self.main_keys:
			try:
				row = col_a.index(key)+1
			except ValueError:  #the section doesnt exist in the tab
				continue
			boundary_rows.append(row)
		if boundary_rows:
			last_index = boundary_rows[-1]
			lastrow = sheet.range(f"A{last_index}").current_region.rows.count + last_index + 1
			boundary_rows.append(lastrow)
		return boundary_rows


	def get_unique_names(self):
		"""
		Gets all the unique names across all the tabs. Searches for all names within any 2 main keys.
		This assumes there is always 2 empty cells between the end of a section and start of the next.

		Also cleans the tabs list to remove tabs that are not of interest
		"""
		names = []
		sorted_tabs = []
		for sheet_name in self.tabs:
			sht = self.book.sheets[sheet_name]
			boundary_rows = self.get_section_boundaries(sheet=sht)
			total = len(boundary_rows)-1
			if boundary_rows: # boundary rows list empty => none of the sections exist in tab
				state_name = sht.range('A5').value
				if state_name in [None, ""]:
					self.WshShell.Popup(f"Sheet [{sheet_name}] is missing a state in cell A5. Please fill and try again.", 10, "Missing info", 0+16)  #0=OK button, 16=critical
					return None,None
				sorted_tabs.append(f"{sht.range('A5').value}-{sheet_name}")   #Combines state and tab name. Used for sorting the tab columns
				self.boundary_dict[sheet_name] = boundary_rows
				for i in range(0,total):
					section_start = boundary_rows[i]
					section_end = boundary_rows[i+1] - 3
					names.extend(sht.range((section_start,2),(section_end,2)).value)

		names = set(filter(None,names))   # None entries arise due to unexpected blank spaces within the sections 

		return names,sorted_tabs


	def create_main_dict(self,sheet,boundary_rows) -> dict:
		"""
		Creates a main dictionary of dictionaries mapping each section's {name:formula}
		e.g. {'Allocation of Taxable Income (Loss) per Partners': {
						'Rest Illinois Holdco, LLC': "='Cidge Care & Rehab'!$D$78",
						'MRS Family Trust': "='Cidge Care & Rehab'!$D$79",
						'Yehuda Stern': "='Cidge Care & Rehab'!$D$86",
						}
			'Allocation of Ordinary Income (Loss) per Partners': {None: "='Cidge Care & Rehab'!$D$94",
						'Crest Illinois Holdco, LLC': "='Cidge Care & Rehab'!$D$95",
						'MRS Family Trust': "='Cidge Care & Rehab'!$D$96",
						},
						........
		"""

		main_dict = {}
		for i in range(0,len(boundary_rows)-1):
			row = boundary_rows[i]
			endrow = sheet.range(f"A{row}").end('down').row - 1   #end of a section
			cell_ranges = sheet.range(f"D{row+2}:D{endrow}")    #row+2 -> assumes theres always a blank space between section title and the first name       
			values = [self.create_address(rng) for rng in cell_ranges]   #Derive cell addresses from range names.
			dicti = dict(zip(sheet[f"B{row+2}:B{endrow}"].value,values))   ##Map name with address
			section = sheet.range(f"A{row}").value
			main_dict[section] = dicti

		return main_dict


	def create_address(self,cell_range:range) -> str:
		"""
		Example cell_range: <Range [Sample Summary filled by hand.xlsx]Cidge Care & Rehab!$D$95>
		"""
		sheet_name = cell_range.sheet.name
		cell = cell_range.address
		formula = f"='{sheet_name}'!{cell}"
		return formula


	def finalise(self,tab_names,state_names,names, data_list, boundary_columns) -> None:
		"""
		Adds the data to be spreadsheet and does some minimal styling
		"""	
		summary_endrow = 11 + len(data_list[0]) - 1  #11 is row with first name
		summary_endcolumn = 5 + len(tab_names)-1

		stated_items = [
					"ORDINARY",
					"GUARANTEED PYMT",
					"RENTAL",
					"CAPITAL GAIN (LOSS)",
					None   
				]

		name_output = []
		rw = []
		j = 0
		for name in names:
			row = 11 + j
			name_output.extend([name,None,None,None,None])
			j+=5      #spacing of 5 btwn names

			#Create ranges to be styled. rw = rows
			rw.append(f"A{row}:A{row+3}")

		#Paste names
		self.summary_sht.range(11,1).options(transpose=True).value = name_output
		self.summary_sht.range(11,3).options(transpose=True).value = stated_items*len(names)

		#Paste headings
		self.summary_sht.range("E6").value = state_names
		self.summary_sht.range("E7").value = tab_names

		#Paste main data/formulas
		self.summary_sht.range((11,5),(summary_endcolumn,summary_endrow)).options(transpose=True).value = data_list

		#Style data section
		self.toggle_app_state(val=False)  #turn off screen updating
		self.add_mainstyle(rows=rw,endcol=summary_endcolumn)

		#populate total formulas
		self.add_total_formulas(boundary_columns,endrow=summary_endrow)

		#Remove styling from unused cells
		self.deformat_cells(start_col=summary_endcolumn+1)


	def add_mainstyle(self,rows:list,endcol:int) -> None:
		"""
		Add merging and color to the data section
		"""
		#Style rows
		slices = helpers.slice_ranges(rows,20)
		for rng in slices:
			self.summary_sht.range(self.summary_sht.range(rng)).color = 14998742    #Color a cell. Color code got from vba using Activecell.Interior.Color
		
		#Style columns
		for i in range(1,endcol,2):
			next_col = self.get_column_letter(i,offset=2)  #first column picked is C
			new_slices = [s.replace("A", next_col) for s in slices]  #pattern is the same so simply replace column A ranges with new column letters
			for rng in new_slices:
				self.summary_sht.range(self.summary_sht.range(rng)).color = 14998742

		
	def add_header_styles(self,startcol,endcol) -> None:
		"""
		Adds colors and borders to header section
		"""
		s =  self.summary_sht.range((6,startcol),(6,endcol))
		s.merge()
		s.api.BorderAround(11,-4138)  #BorderAround(LineStyle,Weight) - values from MS docs

		s =  self.summary_sht.range((7,startcol),(9,endcol))
		s.api.BorderAround(11,-4138)
		s.color = 14998742


	def deformat_cells(self,start_col) -> None:
		"""
		Removes default styling in unused columns. 
		300 columns were prestyled in template (rows 7-9)
		"""
		s = self.summary_sht.range((7,start_col),(9,500))
		s.unmerge()
		s.api.VerticalAlignment = -4108   #Centre alignment
		s.api.WrapText = False
		s.api.ColumnWidth = 13.29


	def add_total_formulas(self,boundary_columns,endrow):
		"""
		Boundary columns holds start and end positions of the states e.g. [[5, 9],[11, 19],[21, 27]],
		Formula is created is 4 cells (row 11-14) then autofilled to the last row
		Header section styling also done here
		"""
		for boundary in boundary_columns:
			startcol = boundary[0]
			endcol = boundary[1]
			rg = str(self.summary_sht.range((11,startcol),(11,endcol-2)).address).replace('$','')
			formula = f"=SUM({rg})"
			
			self.summary_sht.range((11,endcol)).formula = formula
			self.summary_sht.range((11,endcol),(11+3,endcol)).api.FillDown()

			#autofill to the end, leaving a space every 5th column
			totals_col = self.get_column_letter(num=endcol,offset=0)
			fillrange = f"{totals_col}11:{totals_col}{endrow}"
			self.summary_sht.range((11,endcol),(15,endcol)).api.AutoFill(self.summary_sht.range(fillrange).api,AutoFillType.xlFillDefault)

			#Style header section
			self.add_header_styles(startcol=startcol,endcol=endcol)
		
	def get_column_letter(self,num,offset):
		"""
		Converts column number to a letter e.g. 10 = J
		"""
		letters = 2
		if num>24:
			letters = 3  #allows picking 2 letters when columns change from single to double eg Z to AA
		addr = self.summary_sht.range(1,num)
		letter = addr.offset(0,offset).address[1:letters]
		return letter


	def toggle_app_state(self,val:bool):
		self.excelapp.ScreenUpdating = val
		self.excelapp.DisplayAlerts = val


def make_summaries():
	mainbook = xw.Book.caller()
	x = SummaryProcessor(workbook=mainbook)
	x.make_summaries()

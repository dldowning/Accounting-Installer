import xlwings as xw
from xlwings.constants import HAlign, VAlign, BordersIndex, LineStyle
import win32com.client as winclient
import helpers
import constants

class IndividualSummaryProcessor:
	def __init__(self,workbook) -> None:
		self.book = workbook
		self.summary_sht = workbook.sheets("Summary")
		self.companies = []
		self.company_columns = []
		self.shareholders = []
		self.shareholder_rows = []
		self.lastcol = self.summary_sht.range("ZZ6").end('left').column
		self.states_data = self.summary_sht.range((6,5),(6,self.lastcol)).value
		self.state_indices = helpers.get_indices(data = self.states_data, correction = 5)
		self.main_items = constants.main_items
		self.item_totals = constants.item_totals
		self.final_totals = constants.final_totals
		self.final_notes = constants.final_notes
		
	WshShell = winclient.DispatchEx("WScript.Shell")
	def generate_individual_summaries(self):
		"""
		Generates summaries for each owner. Creates new tabs to the right for each individual owner,
		with all the company info they own a percentage for.
		
		"""		
		#Get shareholders. 
		self.shareholders, self.shareholder_rows = self.get_shareholders_info()
		
		#Get companies info
		self.companies, self.company_columns = self.get_companies()

		#Create companies dict
		company_dict = self.create_company_dict()

		#Create shareholders dict
		lastrow = self.shareholder_rows[-1] + 3
		shareholder_dict = self.create_shareholder_dict(lastrow)
		
		for owner in self.shareholders:
			#Copy Template to owner individual summary
			try:
				self.book.sheets('Individual Summary Template').copy(name=owner)
			except ValueError:
				self.WshShell.Popup(f"Sheet named {owner} already exists > Exiting now..", 10, "Duplicate", 0+16)  #0=OK button, 16=critical, 10=on for 10 seconds
				return

			owner_sht = self.book.sheets(owner)
			owner_sht.range('A1:A6').clear_contents()
			owner_sht.range('B7:D1000').clear()


			all_headings, all_formulas, all_stated_item_rows, bonus_formula = self.create_summary_layout(owner,shareholder_dict,company_dict,owner_sht)
			total_ordinary, total_guaranteed, total_rental, total_capital, summary_formulas = helpers.create_summary_formulas(all_formulas,all_stated_item_rows)
			
			#Add top titles
			owner_sht.range("A1").value = owner

			#Populate headings
			all_headings.extend(constants.final_totals)
			all_headings.append(None)
			all_headings.extend(constants.final_notes)
			owner_sht.range("C7").options(transpose=True).value = all_headings

			#Populate formulas
			owner_sht.range("D8").options(transpose=True).value = summary_formulas

			#Style and finalise layout
			end, state_total_rows = self.finalise_layout(all_stated_item_rows, owner_sht)

			#End table formulas
			taxable_income = ''
			for each in state_total_rows:
				taxable_income = taxable_income + f"+D{each}" 
			tf = [taxable_income,total_ordinary,total_guaranteed,total_rental,total_capital,bonus_formula]
			table_formulas = [f.replace(f[0],'=',1) for f in tf]
			owner_sht.range(end+7,4).options(transpose=True).value = table_formulas


	def get_state(self, company_column):
		"""
		Gets state name based on company column number
		"""
		col_number =  min(self.state_indices, key=lambda x:abs(x-company_column))
		state = self.states_data[col_number - 5]
		
		return state

	def get_state_info(self, states_data, company_column, correction):
		"""
		Gets state name based on company column number
		correction: Used to calculate the actual row/column numbers from the indices
		"""
		state_indices = helpers.get_indices(data = states_data, correction = correction)
		col_number =  min(state_indices, key=lambda x:abs(x-company_column))
		state = states_data[col_number - correction]

		return state,state_indices

	def create_company_dict(self):
		company_dict = {}
		for i, col in enumerate(self.company_columns):
			company = self.companies[i]
			state = self.get_state(col)
			company_dict[company] = {}
			company_dict[company]['State'] = state
			company_dict[company]['Column'] = col
		return company_dict

	def get_companies(self):
		lastcol = self.summary_sht.range("ZZ7").end('left').column
		companies_data = self.summary_sht.range((7,5),(7,lastcol)).value
		companies = []
		for each in companies_data:
			if each == None or "Total" in each:
				continue
			companies.append(each)

		#Get start end columns of each state. Add 5 as first column is E
		company_columns = helpers.get_indices(data = companies_data, correction = 5, elements = companies)
		
		return companies, company_columns

	def get_shareholders_info(self):
		"""
		Get list shareholders and the row number. 
		To find last shareholder, search last 'Capital Gain' heading and use it to identify that.
		"""
		stated_items = self.summary_sht.range("C1:C500").value
		stated_items.reverse()
		last_capitalgain = stated_items.index('CAPITAL GAIN (LOSS)')
		lastrow = len(stated_items) - last_capitalgain - 3
		shareholder_data = self.summary_sht.range(f"A11:A{lastrow}").value
		shareholders = [each for each in shareholder_data if each != None]
		shareholder_rows = helpers.get_indices(data = shareholder_data, correction=11, elements = shareholders)
		return shareholders, shareholder_rows

	def create_shareholder_dict(self, lastrow):
		"""
		Creates a dictionary containing all owner row number and list of companies they own: 
		{
			Owner: {
				Index: 2,
				Companies: [Company1, Company2,...]
			},
			....
		}
		"""
		shareholder_dict = {}
		for count, col in enumerate(self.company_columns):
			company_entries = self.summary_sht.range((1,col),(lastrow,col)).formula
			company_entries = [item for sublist in company_entries for item in sublist]
			company = self.companies[count]

			for i in range(0,len(self.shareholder_rows)):
				index = self.shareholder_rows[i]
				try:
					next_index = self.shareholder_rows[i+1]
				except IndexError:
					next_index = index + 4
				shareholder_name = self.shareholders[i]

				if count == 0:
					shareholder_dict[shareholder_name] = {
														'Index':index,
														'Companies':[],
															}

				stated_items = company_entries[index:next_index-1]
				if stated_items.count('') < 4:
						shareholder_dict[shareholder_name]['Companies'].append(company)
					
		return shareholder_dict
		

	def create_summary_layout(self,owner, shareholder_dict, company_dict,owner_sht):
		"""
		Creates the layout and formulas for each shareholder
		"""
		start_row = 8    #Stated item entries start at row 8
		all_headings = []
		all_formulas = []
		all_stated_item_rows = []

		# for owner in self.shareholders[15:16]:
		shareholder_row = shareholder_dict[owner]['Index']
		companies = shareholder_dict[owner]['Companies']
		company_states = helpers.create_state_dict(companies,company_dict)
		bonus_formula = ''
		for state in company_states:
			owner_companies = company_states[state]['Companies']
			total_companies = len(owner_companies)
			is_last = False
			temp = []
			y=[]
			state_formulas = []
			owner_sht.range(start_row-1,2).value = state
			for i, company in enumerate(owner_companies):
				company_column = company_dict[company]['Column']
				if i == total_companies - 1:
					is_last = True

				headings = helpers.create_item_headings(company, state, is_last)
				if is_last:
					headings.append(f"TOTAL {state.upper()}")

				all_headings.extend(headings)
				all_headings.append(None)
				stated_item_rows = [i for i in range(start_row, start_row+4)]
				temp.append(stated_item_rows)
				start_row+=6
				
				column_letter = self.get_column_letter(company_column,offset=0)
				formulas = helpers.create_formulas(shareholder_row, column_letter, is_last)
				state_formulas.extend(formulas)
				y.append(stated_item_rows)
			
			#Get depreciation bonus
			bonus_formula = bonus_formula + "+" + self.get_bonus_formula(shareholder = owner, sheet_name = company)
			
			all_formulas.append(state_formulas)
			all_stated_item_rows.append(y)
			start_row+=5

		return all_headings, all_formulas, all_stated_item_rows, bonus_formula

	def finalise_layout(self, all_stated_item_rows, owner_sht):
		"""
		Adds styling and creates formulas for the last table
		"""
		state_total_rows = []
		for each in all_stated_item_rows:
			for i, rows in enumerate(each):
				if i == 0:
					start_row = min(rows)-1

				start = min(rows)
				end = max(rows)
				owner_sht.range(start-1,3).font.bold = True              #Bold heading
				owner_sht.range((start,3),(end,3)).api.IndentLevel = 3   #indent stated items
				if i == len(each)-1:
					state_total_rows.append(end+5)

					#Style and right align totals
					owner_sht.range((end+1,3),(end+5,3)).font.bold = True
					owner_sht.range((end+1,3),(end+5,3)).api.HorizontalAlignment = HAlign.xlHAlignRight

					#Section border
					s = owner_sht.range((start_row,3),(end+4,4))
					s.api.BorderAround(11,2)  #BorderAround(LineStyle,Weight) - values from MS docs
					b = owner_sht.range((start_row,2),(end+4,2))
					b.api.BorderAround(11,2)
					owner_sht.range(end,4).api.Borders(BordersIndex.xlEdgeBottom).LineStyle = LineStyle.xlContinuous

					#Style state cell
					b.merge()
					b.api.VerticalAlignment = VAlign.xlVAlignCenter
					b.api.HorizontalAlignment = HAlign.xlHAlignCenter
					b.font.bold = True
					owner_sht.range(start_row,2).api.Orientation = 90

		#Style last table
		s = owner_sht.range((end+7,3),(end+16,4))
		s.api.BorderAround(11,2)  #BorderAround(LineStyle,Weight) - values from MS docs
		s.font.bold = True
		s.api.Borders(BordersIndex.xlInsideHorizontal).LineStyle = LineStyle.xlContinuous

		#Style notes
		s = owner_sht.range((end+17,3),(end+22,3))
		s.font.size = 10
		s.api.WrapText = True
		s.font.italic = True
		s.offset(1).font.color = (255,0,0)   #notes to red
		owner_sht.range(end+23,3).font.bold = True

		#White background
		owner_sht.range((1,1),(end+27,5)).color = (255,255,255)

		#Merge top section
		owner_sht.range('A1:E1').merge()
		owner_sht.range('A2:E2').merge()
		owner_sht.range('A3:E3').merge()
		owner_sht.range('A4:E4').merge()

		return end, state_total_rows
    
	#TAXPAYER TOTAL FEDERAL BONUS DEPRECIATION formula
	def get_bonus_formula(self,shareholder, sheet_name):
		"""
		Searches the specified company spreadsheet for the owner name and links to the bonus cell
		e.g. 'Fort Indiana Holdco'!D170
		"""
		sheet = self.book.sheets(sheet_name)
		d = sheet.range("A1:A10000").value
		idx = d.index('Allocation of Federal Bonus Depreciation per Partners') + 1
		d = sheet.range("B1:B1000").value
		bonus_row = d[idx:].index(shareholder) + idx + 1
		formula = f"'{sheet_name}'!D{bonus_row}"
		
		return formula
		
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


def generate_individual_summaries():
	mainbook = xw.Book.caller()
	x = IndividualSummaryProcessor(workbook=mainbook)
	x.generate_individual_summaries()

mainbook = xw.Book("Test Sample.xlsx")
x = IndividualSummaryProcessor(workbook=mainbook)
x.generate_individual_summaries()
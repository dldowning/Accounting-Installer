import constants
import win32api
import subprocess

def add_spacers(listt,spacer) -> list:
		"""
		For adding column spacing after every tab name
		"""
		i = 1 
		while i < len(listt):
			listt.insert(i,spacer)
			i = i + 2
		return listt

def create_final_list(names,dictionary,tabs,main_keys,states) -> list:
	"""
	Create the complete list with  all the tabs data
	"""
	data_list = []
	current_state = states[0]   #states list is synced to tabs list ie states[0] corresponds to tabs[0]
	final_headings = []
	final_states = []
	boundary_columns = []
	boundary = 3
	list_total = len(names)*5
	k = 0
	for i,sheet_name in enumerate(tabs):
		state_name = states[i]
		tab_data = []
		for name in names:
			listt = []
			for key in main_keys:
				try:
					listt.append(dictionary[sheet_name]['main'][key][name]) 
				except KeyError:
					listt.append(None)
			listt.append(None)
			tab_data.extend(listt)

		#insert totals column
		if state_name !=current_state:
			final_headings.extend([f"Total {states[i-1]}",None])
			final_states.extend([None,None])
			data_list.append([None]*list_total)   #spacer for total
			data_list.append([None]*list_total)   #TO INSERT TOTAL FORMULAS HERE
			current_state = states[i]
			boundary_columns.append([boundary+2,(i*2)+5+k])   #formula derived from analysing column patterns
			boundary = (i*2)+5+k
			k+=2
		final_headings.extend([sheet_name,None])  #creates spacing between tab headings
		final_states.extend([state_name,None])  

		data_list.append(tab_data)
		data_list.append([None]*list_total)   #space after every column

	boundary_columns.append([boundary+2,(i*2)+7+k])
	final_headings.append(f"Total {states[i-1]}")

	return data_list,final_headings,final_states,boundary_columns

def sort_tabs(sorted_tabs) -> list:
	#Sort the tabs in order of state
	sorted_tabs = sorted(sorted_tabs)
	tabs_output = [each.split('-')[1] for each in sorted_tabs]
	sorted_states = [each.split('-')[0] for each in sorted_tabs]
	return tabs_output,sorted_states

def slice_ranges(rng_list,i) -> list:
	"""
	The range function has a 250 character limit so have to slice large string ranges to 
	stay within limit by separating to smaller groups
	For single ranges e.g. $A$11,$A$16,... limit to 20 cells, 
	For double e.g. 'A11:A14,A16:A19,A21:A24,... limit to 10
	Returns: list of lists of grouped ranges
	"""
	sliced = []
	for n in range(0, len(rng_list),i):
		rng = rng_list[n:n + i]
		output = ','.join(rng)

		sliced.append(output)

	return sliced

def get_address_rows(address,column,pattern) -> list:
	"""
	Change cell address format to R1C1 format eg B5 to (2,5)
	"""
	startrow,endrow = pattern.findall(address)
	cell_address = ((int(startrow),column),(int(endrow),column))
	return cell_address

def get_indices(data, correction = 0, elements = None):
	"""
	Gets the position (rows or columns) of elements passed to it ,
	correction: Used to calculate the actual row/column numbers from the indices
	"""
	if elements == None:
		elements = [each for each in data if each != None]
		
	indices = [data.index(element) + correction for element in elements]
	
	return indices

def get_state_info(states_data, company_column, correction):
	"""
	Gets state name based on company column number
	correction: Used to calculate the actual row/column numbers from the indices
	"""
	state_indices = get_indices(data = states_data, correction = correction)
	col_number =  min(state_indices, key=lambda x:abs(x-company_column))
	state = states_data[col_number - correction]

	return state,state_indices

def create_formulas(shareholder_row,company_column, is_last=False):
	"""
	Creates the formulas that link to each owned companies' stated item values.
	There's 4 entries Ordinary, Fuarantted Pymt, Rental and Capital Gain.
	e.g ='Summary Sample'!Q91
	"""
	formulas = [f"='Summary'!{company_column}{i}" for i in range(shareholder_row,shareholder_row+4)]
	
	if is_last == True:
		return formulas
	
	formulas.extend(["",""])    #Leave blank space if there is multiple companies for the state
	
	return formulas

def create_item_headings(company_name, state,  is_last):
	headings = [each.replace('{state}',state) for each in constants.main_items]
	headings.insert(0,company_name)
	
	if is_last:
		t = [each.replace('{state}',state) for each in constants.item_totals]
		headings.extend(t)
	return headings
	
def create_state_dict(companies,company_dict):
	"""
	Creates dictionary of state against its list of comapnies.
	e.g. {Indiana:[Company1,Company2..], Illinois:{Company1,Company3,...}}
	"""
	state_dict = {}
	for company in companies:
		state = company_dict[company]['State']
		if state not in state_dict:
			state_dict[state] = {'Companies':[]}
		state_dict[state]['Companies'].append(company)
		
	return state_dict

def create_item_total(state_items,start,end):
	state_items_flattened = [item for sublist in state_items for item in sublist[start:end]]
	first = f"=D{state_items_flattened[0]}"
	for i in state_items_flattened[1:]:
		first += "+D%s"%(i) 
	return first



def create_summary_formulas(all_formulas,all_stated_item_rows):
	"""
	Links the stated items to the Summary tab
	e.g. ='Summary'!E86
	"""
	summary_formulas = []
	total_ordinary = ''
	total_guaranteed = ''
	total_rental = ''
	total_capital = ''
	for i,each in enumerate(all_stated_item_rows):
		if len(all_formulas[i])>1:
			summary_formulas.extend(all_formulas[i])
		else:
			summary_formulas.append(all_formulas[i])
		all_items = [item for sublist in each for item in sublist]
		total_formula_end = max(all_items)
		
		#Ordinary
		ordinary = create_item_total(each,start = 0, end = 1)
		summary_formulas.append(ordinary)
		total_ordinary+=f"+{ordinary[1:]}"
		
		#Guaranteed
		guaranteed = create_item_total(each,start = 1, end = 2)
		summary_formulas.append(guaranteed)
		total_guaranteed+=f"+{guaranteed[1:]}"
		
		#Rental
		rental = create_item_total(each,start = 2, end = 3)
		summary_formulas.append(rental)
		total_rental+=f"+{rental[1:]}"
		
		#Capital
		capital = create_item_total(each,start = 3, end = 4)
		summary_formulas.append(capital)
		total_capital+=f"+{capital[1:]}"
		
		#State total
		state_total = f"=SUM(D{total_formula_end+1}:D{total_formula_end+4})"
		summary_formulas.extend([state_total,"",""])
	
	return total_ordinary, total_guaranteed, total_rental, total_capital, summary_formulas

def pdfmail(filepath):
	#Open default mail client
	win32api.ShellExecute(0,'open','mailto:',None,None ,0)

	#Open and highlight file to attach
	subprocess.Popen(f'explorer /select,"{filepath}"')

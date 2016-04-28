

'''Downloads the corporate spreadsheet'''
gConnection = GoogleConnection()
gConnection.ConnectToGoogle()
gConnection.DownloadCorporateData()

'''Loads the downloaded file'''
corporateDataReader = csv.reader(open(gConnection.corporateFilePath,'r'), delimiter=',')

#The dictionary of company contact info keyed by company
contact_sheet_by_company = dict()

#spreadsheet index constants
company_name_index = 1
contact_name_index = 2
contact_email_index = 3
contact_phone_index = 4
contact_address_index = 5
contact_type_index = 6
known_by_index = 7 
notes_index = 8

#dictionary keys
contact_name = 'Contact_Name'
contact_email = 'Email'
contact_phone = 'Phone'
contact_address = 'Address'
contact_type = 'Type'
known_by = 'Member_Known_By'
contact_notes = 'Notes'

corporateDataReader.next()

#Gets the string value of a cell
def get_spreadsheet_cell(row,col):
	return str(s.cell(row,col).value)

s=wb.sheet_by_index(0)
#The dictionary of company contact info keyed by company
contact_sheet_by_company = dict()

#spreadsheet index constants
company_name_index = 1
contact_name_index = 2
contact_email_index = 3
contact_phone_index = 4
contact_address_index = 5
contact_type_index = 6
known_by_index = 7 
notes_index = 8

#dictionary keys
contact_name = 'Contact_Name'
contact_email = 'Email'
contact_phone = 'Phone'
contact_address = 'Address'
contact_type = 'Type'
known_by = 'Member_Known_By'
contact_notes = 'Notes'

for row in range(s.nrows):
	company_name=get_spreadsheet_cell(row,company_name_index)
	if company_name not in contact_sheet_by_company:
		contact_sheet_by_company[company_name]=[]
	
	contact_info = dict()
	contact_info[contact_name]=get_spreadsheet_cell(row,contact_name_index)
	contact_info[contact_email]=get_spreadsheet_cell(row,contact_email_index)
	contact_info[contact_phone]=get_spreadsheet_cell(row,contact_phone_index)
	contact_info[contact_address]=get_spreadsheet_cell(row,contact_address_index)
	contact_info[contact_type]=get_spreadsheet_cell(row,contact_type_index)
	contact_info[known_by]=get_spreadsheet_cell(row,known_by_index)
	contact_info[contact_notes]=get_spreadsheet_cell(row,notes_index)
	contact_sheet_by_company[company_name].append(contact_info)
	
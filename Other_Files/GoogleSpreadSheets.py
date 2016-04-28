import gdata.docs.data
import gdata.docs.client
import gdata.docs.service
import gdata.docs
import gdata
import getpass
class GoogleConnection:
	client=None;
	electees='Electee Group Formation Survey'
	electeeFilePath=r'''C:\Users\Mike\Dropbox\TBP\Electees\ElecteeData.csv'''
	corporateSheetName='Tau Beta Pi Corporate Contacts'
	corporateFilePath=r'''C:\Users\Mike\Dropbox\TBP\CorporateRelations\CorporateData.csv'''
	def ConnectToGoogle(self):
		self.client = gdata.docs.client.DocsClient()
		self.client.ssl = True
		self.client.http_client.debug = False
		username = raw_input('Google username: ')
		password = getpass.getpass('Password: ')
		self.client.ClientLogin(username,password,self.client.source)

	def DownloadElecteeData(self):	
		self.DownloadSpreadsheet(self.electees,self.electeeFilePath)

	def DownloadCorporateData(self):
		self.DownloadSpreadsheet(selfcorporateSheetName,self.corporateFilePath)
	def DownloadSpreadsheet(self,name,toLocation):

		feed = self.client.GetAllResources()
		for  document_entry in feed:
			if document_entry.title.text == name:
				print 'match found'
				self.client.DownloadResource(entry=document_entry,file_path=toLocation,extra_params={'gid': 0, 'exportFormat': 'csv'})
			
'''This does update a spreadsheet just screws up form stuff'''		
#client.UpdateResource(entry=document_entry,media=ms,update_metadata=False)
#ms=gdata.data.MediaSource(file_path=change_file_path, content_type='application/vnd.ms-excel')

#extra_params={'gid': 0, 'exportFormat': 'csv'}
#PrintFeed(feed)
#entry = feed.entry[0]
#new_sheet= gdata.data.MediaSource(file_path=r'''C:\Users\Mike\Documents\REPLACE.xls''',content_type=gdata.docs.service.SUPPORTED_FILETYPES['XLS'])
#updated_entry = client.Put(new_sheet,entry[0].GetEditMediaLink().href)
#PrintFeed(feed)


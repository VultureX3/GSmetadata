import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools


class Project:

	def __init__(self, metadata_GS_ID, admin_email, info_file):
		self.project_name = ''
		self.metadata_GS_ID = metadata_GS_ID
		self.folder = None
		self.admin_email = admin_email
		self.info_file = info_file
		self.client = None
		self.service = None
		self.metadata = {}

	def __get_metadata(self):
		metadata_GS = self.client.open_by_key(metadata_GS_ID)
		for sheet in metadata_GS:
			self.metadata[sheet.title] = sheet.get_all_records()
		self.project_name = self.metadata['information'][0]['Значение']

	def __create_folder(self):
		folder_metadata = {
    		'name': self.project_name,
    		'mimeType': 'application/vnd.google-apps.folder'
		}
		self.folder = self.service.files().create(body=folder_metadata,
                                    fields='id').execute()
		print(self.folder['id'])

	def __create_accounts(self):
		for developer in self.metadata['participants']:
		    spreadsheet = self.client.create('Учет трудозатрат %s проект %s' % 
		    			(developer['Имя'], self.project_name))
		    self.client.insert_permission(
		    						file_id=spreadsheet.id, 
		    						value=admin_email,
		    						perm_type='user',
		    						role='owner',
		    						notify=False,
		    						email_message=None
		    						)
		    try:
			    self.client.insert_permission(
			    						file_id=spreadsheet.id, 
			    						value=developer['email'],
			    						perm_type='user',
			    						role='writer',
			    						notify=False,
			    						email_message=None
			    						)
		    except:
			    print('No Google user with email %s' % developer['email'])
		    worksheet = spreadsheet.add_worksheet('timesheet', 300, 12)
		    spreadsheet.del_worksheet(spreadsheet.sheet1)
		    for cell, text in zip (('A1', 'B1', 'C1', 'D1'),
		    		('Дата', 'Название работ', 'Объем работ, часов', 'Статус оплаты (заполняется только РП)')):
		    	worksheet.update_acell(cell, text)
		    self.service.files().update(fileId=spreadsheet.id,
                                    addParents=self.folder['id'],
                                    fields='id, parents').execute()


	def main(self):
		scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
		credentials = ServiceAccountCredentials.from_json_keyfile_name(self.info_file, scope)
		self.client = gspread.authorize(credentials)

		store = file.Storage('token.json')
		creds = store.get()
		if not creds or creds.invalid:
			flow = client.flow_from_clientsecrets('credentials.json', scope)
			creds = tools.run_flow(flow, store)
		self.service = build('drive', 'v3', http=creds.authorize(Http()))
		print(dir(self.service))

		self.__get_metadata()
		self.__create_folder()
		self.__create_accounts()


if __name__ == '__main__':
	metadata_GS_ID = '1eCVqx5GhgY_LXcYcLIcC0_2Rw4cAi8Elyx7i9o3l5ko'
	folder_ID = '1fhgm1Rpz3Wv_CL1hZzmd24sZKn-UJj6q'
	admin_email = 'ikhalepsky@gmail.com'
	info_file = 'Metadata-257dae3670d7.json'
	my_project = Project(metadata_GS_ID, admin_email, info_file)
	my_project.main()

import gspread
from oauth2client.service_account import ServiceAccountCredentials


class Project:

	def __init__(self, metadata_GS_ID, folder_ID, admin_email, info_file, dev_emails):
		self.metadata_GS_ID = metadata_GS_ID
		self.folder_ID = folder_ID
		self.admin_email = admin_email
		self.info_file = info_file
		self.dev_emails = dev_emails
		self.client = None
		self.metadata = {}

	def __get_metadata(self):
		metadata_GS = self.client.open_by_key(metadata_GS_ID)
		# print(dir(metadata_GS.sheet1))
		for sheet in metadata_GS:
			self.metadata[sheet.title] = sheet.get_all_records()

	def __create_accounts(self):
		for developer in self.metadata['participants']:
			print(developer)
		    # sheet = self.client.create('Учет трудозатрат %s проект %s' % 
		    # 			(developer['Имя'], self.metadata['information'][0]['Значение']))
		    # self.client.insert_permission(
		    # 						file_id=sheet.id, 
		    # 						value=admin_email,
		    # 						perm_type='user',
		    # 						role='owner',
		    # 						notify=False,
		    # 						email_message=None
		    # 						)
		    # print(sheet.id)
	    # if parent_folder_ids:
	    # 	body["parents"] = folder_ID


	def main(self):
		scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
		credentials = ServiceAccountCredentials.from_json_keyfile_name(self.info_file, scope)
		self.client = gspread.authorize(credentials)
		self.__get_metadata()
		self.__create_accounts()


if __name__ == '__main__':
	metadata_GS_ID = '1eCVqx5GhgY_LXcYcLIcC0_2Rw4cAi8Elyx7i9o3l5ko'
	folder_ID = '1fhgm1Rpz3Wv_CL1hZzmd24sZKn-UJj6q'
	admin_email = 'ikhalepsky@gmail.com'
	info_file = 'GSmetadata-b4b4a0eaa7d3.json'
	my_project = Project(metadata_GS_ID, folder_ID, admin_email, info_file, [])
	my_project.main()

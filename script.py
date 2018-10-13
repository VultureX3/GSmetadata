import gspread
from oauth2client.service_account import ServiceAccountCredentials

def get_metadata(client, metadata_GS_ID):
	metadata_GS = client.open_by_key(metadata_GS_ID)
	metadata = {}
	# print(dir(metadata_GS.sheet1))
	for sheet in metadata_GS:
		metadata[sheet.title] = sheet.get_all_records()
	return metadata

def create_accounts(client, admin_email, folder_ID, metadata):
	for developer in metadata['participants']:
	    sheet = client.create('Учет трудозатрат %s проект %s' % 
	    			(developer['Имя'], metadata['information'][0]['Значение']))
	    client.insert_permission(
	    						file_id=sheet.id, 
	    						value=admin_email,
	    						perm_type='user',
	    						role='owner',
	    						notify=False,
	    						email_message=None
	    						)
	    print(sheet.id)
    # if parent_folder_ids:
    # 	body["parents"] = folder_ID


def main():
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	metadata_GS_ID = '1eCVqx5GhgY_LXcYcLIcC0_2Rw4cAi8Elyx7i9o3l5ko'
	folder_ID = '1fhgm1Rpz3Wv_CL1hZzmd24sZKn-UJj6q'
	admin_email = 'ikhalepsky@gmail.com'
	info_file = 'GSmetadata-b4b4a0eaa7d3.json'

	credentials = ServiceAccountCredentials.from_json_keyfile_name(info_file, scope)
	client = gspread.authorize(credentials)
	metadata = get_metadata(client, metadata_GS_ID)
	create_accounts(client, admin_email, folder_ID, metadata)


if __name__ == '__main__':
	main()

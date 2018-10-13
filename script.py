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

    def __move_metadata(self):
        self.service.files().update(fileId=self.metadata_GS_ID,
                                    addParents=self.folder['id'],
                                    fields='id, parents').execute()

    def __create_accounts(self):
        for developer in self.metadata['participants']:
            spreadsheet = self.client.create('Учет трудозатрат %s проект %s' % 
                                             (developer['Имя'], self.project_name))
            self.client.insert_permission(
                                    file_id=spreadsheet.id, 
                                    value=self.admin_email,
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
            for cell, text in zip(('A1', 'B1', 'C1', 'D1'),
                                  ('Дата', 'Название работ', 
                                   'Объем работ, часов', 'Статус оплаты (заполняется только РП)')):
                worksheet.update_acell(cell, text)
            self.service.files().update(fileId=spreadsheet.id,
                                        addParents=self.folder['id'],
                                        fields='id, parents').execute()

    def __create_cost_calc(self):
        spreadsheet = self.client.create('Расчет стоимости проект %s' % self.project_name)
        self.client.insert_permission(
                                    file_id=spreadsheet.id, 
                                    value=self.admin_email,
                                    perm_type='user',
                                    role='owner',
                                    notify=False,
                                    email_message=None
                                    )
        ws_result = spreadsheet.add_worksheet('Итог', 300, 12)
        spreadsheet.del_worksheet(spreadsheet.sheet1)
        for cell, text in zip(('A1', 'B1', 'C1', 'D1', 'E1', 'F1'),
                              ('Исполнитель', 'Должность', 'К оплате, часов', 'Ставка, рублей', 
                               'К оплате, рублей', 'Примечания')):
                ws_result.update_acell(cell, text)
        row = 1
        for developer in self.metadata['participants']:
            row += 1
            for col, value in zip(range(1, 7), (
                    developer['Имя'],
                    developer['Должность'],
                    '''=СУММЕСЛИ('Работы {0}'!$D$2:$D, "Акт", 'Работы {0}'!$C$2:$C)'''.format(developer['Имя']),
                    developer['Ставка внешняя'],
                    '=$C2*$D2',
                    developer['Комментарии к затратам'])):
                ws_result.update_cell(row, col, value)

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

        self.__get_metadata()
        self.__create_folder()
        # self.__move_metadata()
        self.__create_accounts()
        self.__create_cost_calc()


if __name__ == '__main__':
    metadata_GS_ID = '1eCVqx5GhgY_LXcYcLIcC0_2Rw4cAi8Elyx7i9o3l5ko'
    folder_ID = '1fhgm1Rpz3Wv_CL1hZzmd24sZKn-UJj6q'
    admin_email = 'ikhalepsky@gmail.com'
    info_file = 'Metadata-257dae3670d7.json'
    my_project = Project(metadata_GS_ID, admin_email, info_file)
    my_project.main()

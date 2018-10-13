import gspread
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from oauth2client.service_account import ServiceAccountCredentials


class Project:

    def __init__(self, metadata_GS_ID, admin_email, client_email, info_file):
        self.project_name = ''
        self.metadata_GS_ID = metadata_GS_ID
        self.folder = None
        self.admin_email = admin_email
        self.client_email = client_email
        self.info_file = info_file
        self.client = self.service = None
        self.metadata = {}
        self.accounts = {}
        self.cost = self.inner_cost = None

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

    def __create_cost(self):
        spreadsheet1 = self.client.create('Расчет стоимости проект %s' % self.project_name)
        spreadsheet1.share(
                                    value=self.admin_email,
                                    perm_type='user',
                                    role='owner',
                                    notify=False,
                                    email_message=None
                                    )
        self.cost = spreadsheet1

        spreadsheet2 = self.client.create('Внутренние расчеты стоимости проекта %s (конфиденциально)' %
                                        self.project_name)
        spreadsheet2.share(
                                    value=self.admin_email,
                                    perm_type='user',
                                    role='owner',
                                    notify=False,
                                    email_message=None
                                    )
        self.inner_cost = spreadsheet2


    def __move_metadata(self):
        self.service.files().update(fileId=self.metadata_GS_ID,
                                    addParents=self.folder['id'],
                                    fields='id, parents').execute()

    def __create_accounts(self):
        for developer in self.metadata['participants']:
            spreadsheet = self.client.create('Учет трудозатрат %s проект %s' % 
                                             (developer['Имя'], self.project_name))
            spreadsheet.share(
                                    value=self.admin_email,
                                    perm_type='user',
                                    role='owner',
                                    notify=False,
                                    email_message=None
                                    )
            spreadsheet.share(
                                    value=self.client_email,
                                    perm_type='user',
                                    role='writer',
                                    notify=False,
                                    email_message=None
                                    )
            try:
                spreadsheet.share(
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
            self.accounts[developer['Имя']] = spreadsheet

    def __update_cost(self):
        spreadsheet = self.cost
        ws_result = spreadsheet.add_worksheet('Итог', 300, 12)
        spreadsheet.del_worksheet(spreadsheet.sheet1)

        for account in self.accounts:
            ws_work = spreadsheet.add_worksheet('Работы %s' % account, 300, 12)
            ws_work.update_cell(1, 1, '=IMPORTRANGE("%s", "timesheet!A:D")' % 
                                self.accounts[account].id)

        for cell, text in zip(('A1', 'B1', 'C1', 'D1', 'E1', 'F1'),
                              ('Исполнитель', 'Должность', 'К оплате, часов', 'Ставка, рублей', 
                               'К оплате, рублей', 'Примечания')):
                ws_result.update_acell(cell, text)
        row = 2
        for developer in self.metadata['participants']:
            for col, value in zip(range(1, 7), (
                    developer['Имя'],
                    developer['Должность'],
                    '''=СУММЕСЛИ('Работы {0}'!$D$2:$D, "Акт", 'Работы {0}'!$C$2:$C)'''.format(developer['Имя']),
                    developer['Ставка внешняя'],
                    '=$C{0}*$D{0}'.format(str(row)),
                    developer['Комментарии к затратам'])):
                ws_result.update_cell(row, col, value)
            row += 1

        for resource in self.metadata['resources']:
            for col, value in zip((1, 5, 6), (
                    resource['Наименование'],
                    resource['Стоимость, рублей'],
                    resource['Комментарии к затратам'])):
                ws_result.update_cell(row, col, value)
            row += 1
        ws_result.update_cell(row, 1, 'Итого')
        ws_result.update_cell(row, 5, '=СУММ(E2:E%s)' % str(row-1))

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
        self.__create_cost()
        # self.__move_metadata()
        self.__create_accounts()
        self.__update_cost()


if __name__ == '__main__':
    metadata_GS_ID = '1eCVqx5GhgY_LXcYcLIcC0_2Rw4cAi8Elyx7i9o3l5ko'
    folder_ID = '1fhgm1Rpz3Wv_CL1hZzmd24sZKn-UJj6q'
    admin_email = 'ikhalepsky@gmail.com'
    client_email = 'metadata-owner@metadata-1539441882405.iam.gserviceaccount.com'
    info_file = 'Metadata-257dae3670d7.json'
    my_project = Project(metadata_GS_ID, admin_email, client_email, info_file)
    my_project.main()

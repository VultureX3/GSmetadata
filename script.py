import gspread
import json
import re
import time

from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from oauth2client.service_account import ServiceAccountCredentials


class Project:

    def __init__(self, metadata_gs_id, admin_email, info_file, sharing):
        self.project_name = ''
        self.metadata_gs_id = metadata_gs_id
        self.status = self.folder_id = None
        self.admin_email = admin_email
        self.info_file = info_file
        self.sharing = sharing
        self.metadata = {}
        self.accounts = {}
        self.cost_id = self.inner_cost_id = None
        self.__authorize()

    def __authorize(self):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.info_file, scope)
        self.client = gspread.authorize(credentials)

        store = file.Storage('token.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('credentials.json', scope)
            creds = tools.run_flow(flow, store)
        self.service = build('drive', 'v3', http=creds.authorize(Http()))

        metadata_parents = self.service.files().get(fileId=self.metadata_gs_id, fields='parents').execute()
        if len(metadata_parents) > 1:
            self.status = 'update'
            self.folder_id = metadata_parents.get('parents')[1]
        else:
            self.status = 'create'
        if self.status == 'update':
            docs = self.service.files().list(q=f"'{self.folder_id}' in parents").execute()['files']
            self.docs = [doc['name'] for doc in docs]
            for doc in docs:
                if bool(re.match(r'Расчет стоимости*', doc['name'])):
                    self.cost_id = doc['id']
                if bool(re.match(r'Внутренние расчеты*', doc['name'])):
                    self.inner_cost_id = doc['id']

    @staticmethod
    def __share(spreadsheet, value, role):
        sharing = False
        errors = 0
        while not sharing and errors < 3:
            try:
                spreadsheet.share(
                            value=value,
                            perm_type='user',
                            role=role
                            )
                sharing = True
                print(f'{spreadsheet.title}: Permission granted to {value} as {role}')
            except gspread.exceptions.APIError:
                print('Error: ', spreadsheet.title,
                      '''(Don't worry, I'm gonna fix that in a moment, just some Google bugs)''')
                errors += 1

    def read_metadata(self):
        metadata_gs = self.client.open_by_key(self.metadata_gs_id)
        for sheet in metadata_gs:
            self.metadata[sheet.title] = sheet.get_all_records()
        self.project_name = self.metadata['information'][0]['Значение']

    def create_folder(self):
        folder_metadata = {
            'name': self.project_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = self.service.files().create(body=folder_metadata,
                                             fields='id').execute()
        self.folder_id = folder['id']

    def create_cost(self):
        spreadsheet1 = self.client.create(f'Расчет стоимости проект {self.project_name}')
        self.__share(spreadsheet1, self.admin_email, 'owner')
        self.cost_id = spreadsheet1.id

        spreadsheet2 = self.client.create(f'Внутренние расчеты стоимости проекта {self.project_name} (конфиденциально)')
        self.__share(spreadsheet2, self.admin_email, 'owner')

        self.inner_cost_id = spreadsheet2.id

    def create_accounts(self):

        for developer in self.metadata['participants']:
            title = f'Учет трудозатрат {developer["Имя"]} проект {self.project_name}'
            if self.status == 'create' or title not in self.docs:
                spreadsheet = self.client.create(title)
                self.__share(spreadsheet, self.admin_email, 'owner')

                worksheet = spreadsheet.add_worksheet('timesheet', 300, 12)
                spreadsheet.del_worksheet(spreadsheet.sheet1)
                cell_list = worksheet.range('A1:D1')
                for cell, value in zip(cell_list, ('Дата', 'Название работ', 
                                       'Объем работ, часов', 'Статус оплаты (заполняется только РП)')):
                    cell.value = value
                worksheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                if self.sharing:
                    self.__share(spreadsheet, developer['email'], 'writer')

                self.accounts[developer['Имя']] = spreadsheet

    def __create_worksheets(self, spreadsheet):
        for account in self.accounts:
            ws_work = spreadsheet.add_worksheet(f'Работы {account}', 300, 12)
            ws_work.update_cell(1, 1, f'=IMPORTRANGE("{self.accounts[account].id}", "timesheet!A:D")')

    COST_FIELDS = ('Исполнитель', 'Должность', 'К оплате, часов', 'Ставка, рублей', 'К оплате, рублей', 'Примечания')

    def update_cost(self):
        spreadsheet = self.client.open_by_key(self.cost_id)
        if self.status == 'create':
            ws_result = spreadsheet.add_worksheet('Итог', 300, 12)
            spreadsheet.del_worksheet(spreadsheet.sheet1)
        else:
            ws_result = spreadsheet.get_worksheet(0)

        self.__create_worksheets(spreadsheet)

        cell_list = []
        if self.status == 'create':
            cell_list = ws_result.range('A1:F1')
            for cell, value in zip(cell_list, self.COST_FIELDS):
                cell.value = value

        row = 2
        if self.status == 'update':
            row += len(self.metadata['participants']) - len(self.accounts)
        for developer in self.metadata['participants']:
            if developer['Имя'] in self.accounts:
                new_cell_list = ws_result.range(row, 1, row, 6)
                for cell, value in zip(new_cell_list, (
                        developer['Имя'],
                        developer['Должность'],
                        '''=СУММЕСЛИ('Работы {0}'!$D$2:$D, "Акт", 'Работы {0}'!$C$2:$C)'''.format(developer['Имя']),
                        developer['Ставка внешняя'],
                        '=$C{0}*$D{0}'.format(str(row)),
                        developer['Комментарии к затратам'])):
                    cell.value = value
                cell_list += new_cell_list
                row += 1

        new_cell_list = []
        for resource in self.metadata['resources']:
            for col, value in zip((1, 5, 6), (
                    resource['Наименование'],
                    resource['Стоимость, рублей'],
                    resource['Комментарии к затратам'])):
                cell = ws_result.cell(row, col)
                cell.value = value
                new_cell_list.append(cell)
            row += 1

        for col, value in zip((1, 5), (
                    'Итого',
                    f'=СУММ(E2:E{str(row-1)})')):
            cell = ws_result.cell(row, col)
            cell.value = value
            new_cell_list.append(cell)
        cell_list += new_cell_list
        ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')

    INNER_COST_RESULT_HEADERS = (
                                'К оплате заказчиком по актам, рублей',
                                'К оплате разработчикам по актам, рублей',
                                'Фонд встреч, рублей',
                                'Бонусный фонд, рублей',
                                'Предприятию включая операционные расходы и фонд встреч, рублей')

    INNER_COST_HEADERS = ('Основной ФОТ', 'Премиальный ФОТ', 'К оплате', 'Общая задолженность')

    INNER_COST_FIELDS = ('Исполнитель', 'Должность', 'Ставка заказчику, рублей', 'Ставка разработчику, рублей',
                         'Ставка для документов заказчику, рублей', 'По актам, часов', 'Часов для документов',
                         'Стоимость для документов, рублей', 'К оплате разработчику по актам, рублей', '',
                         'Вес вклада, баллы', 'Вклад с учетом веса, баллы', r'Предложение доли фонда, %%',
                         r'Фактическая доля фонда, %%', 'Доля фонда, рублей', '',
                         'К оплате разработчику по актам c премией, рублей', '', 'Всего часов к оплате',
                         'Общая задолженность заказчика, рублей', 'Общая задолженность разработчику, рублей',
                         'Общая задолженность предприятию, рублей')

    @staticmethod
    def __inner_cost_values(developer, row, sum_row):
        values = (developer['Имя'],
                  developer['Должность'],
                  developer['Ставка внешняя'],
                  developer['Ставка внутренняя'],
                  f'=ЕСЛИ($C{row} >= 2000, 2000, $C{row})',
                  '''=СУММЕСЛИ('Работы {0}'!$D$2:$D, "Акт", 'Работы {0}'!$C$2:$C)'''.format(developer['Имя']),
                  f'=$F{row}*($C{row}/$E{row})',
                  f'=$E{row}*$G{row}',
                  f'=$F{row}*$D{row}',
                  '',
                  '',
                  f'=ЕСЛИ($K{row} > 0, ОКРУГЛ($I{row} * $K{row} / $K${sum_row}), 0)',
                  f'=ОКРУГЛ($L{row}*100/$L${sum_row},1)',
                  '',
                  f'=$C$5 * $N{row} / 100',
                  '',
                  f'''=ЕСЛИ(ОКРУГЛ($I{row}+$O{row}, -3) >=$I{row},
                  ОКРУГЛ($I{row}+$O{row}, -3), ОКРУГЛВВЕРХ($I{row}, -3))''',
                  '',
                  '''=СУММЕСЛИ('Работы {0}'!$D$2:$D, "Акт", 'Работы {0}'!$C$2:$C) + СУММЕСЛИ('Работы {0}'!$D$2:$D, "", 'Работы {0}'!$C$2:$C)'''.format(developer['Имя']),
                  f'=$S{row}*$C{row}',
                  f'=$D{row}*$S{row}',
                  f'=$T{row}-$U{row}')
        return values

    def update_inner_cost(self):
        if len(self.accounts) > 15:
            time.sleep(100)

        spreadsheet = self.client.open_by_key(self.inner_cost_id)
        if self.status == 'create':
            ws_result = spreadsheet.add_worksheet('Итог', 300, 26)
            spreadsheet.del_worksheet(spreadsheet.sheet1)
        else:
            ws_result = spreadsheet.get_worksheet(0)

        self.__create_worksheets(spreadsheet)

        cell_list = ws_result.range('A2:A6')
        for cell, value in zip(cell_list, self.INNER_COST_RESULT_HEADERS):
            cell.value = value

        row = 10
        if self.status == 'update':
            row += len(self.metadata['participants']) - len(self.accounts)
        sum_row = row + len(self.accounts)
        fund = ws_result.acell('C4').value
        new_cell_list = ws_result.range('C2:C6')
        for cell, value in zip(new_cell_list, (f'=$H${sum_row}',
                                               f'=$I${sum_row}',
                                               f'{fund}',
                                               '=((($C$2*0.9)-$C$3)/2)-$C$4',
                                               '=$C$2-$C$3-$C$5')):
            cell.value = value
        cell_list += new_cell_list

        new_cell_list = []
        for col, value in zip((3, 11, 17, 19), self.INNER_COST_HEADERS):
            cell = ws_result.cell(8, col)
            cell.value = value
            new_cell_list.append(cell)
        cell_list += new_cell_list

        new_cell_list = ws_result.range('A9:V9')
        for cell, value in zip(new_cell_list, self.INNER_COST_FIELDS):
            cell.value = value
        cell_list += new_cell_list

        row = 10
        if self.status == 'update':
            row += len(self.metadata['participants']) - len(self.accounts)
        sum_row = row + len(self.accounts)
        for developer in self.metadata['participants']:
            if developer['Имя'] in self.accounts:
                new_cell_list = ws_result.range('A{0}:V{0}'.format(str(row)))
                for cell, value in zip(new_cell_list, self.__inner_cost_values(developer, row, sum_row)):
                    cell.value = value
                cell_list += new_cell_list
                row += 1

        cols = ('F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O', 'Q', 'S', 'T', 'U', 'V')
        new_cell_list = [ws_result.acell(f'{col}{row}') for col in cols]
        for cell, col in zip(new_cell_list, cols):
            cell.value = '=СУММ({0}10:{0}{1})'.format(col, str(row-1))
        cell = ws_result.acell(f'A{row}')
        cell.value = 'Итого'
        new_cell_list.append(cell)
        cell_list += new_cell_list
        ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')

    def move_to_folder(self):

        def move(file_id):
            self.service.files().update(fileId=file_id,
                                        addParents=self.folder_id,
                                        fields='id, parents').execute()

        if self.status == 'create':
            for sheet_id in (self.cost_id, self.inner_cost_id, self.metadata_gs_id):
                move(sheet_id)

        for account in self.accounts.values():
            move(account.id)


def main():
    data = json.load(open('admin_data.json'))
    my_project = Project(data['metadata_gs_id'], data['admin_email'], data['info_file'], data['sharing'])
    my_project.read_metadata()
    print('metadata read')
    if my_project.status == 'create':
        my_project.create_folder()
        print('folder created')
        my_project.create_cost()
        print('cost & inner cost sheets created')
    my_project.create_accounts()
    print('accounts created')
    my_project.update_cost()
    my_project.update_inner_cost()
    print('cost & inner cost sheets updated')
    my_project.move_to_folder()
    print('done!')


if __name__ == '__main__':
    main()

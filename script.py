import gspread
import json
import time

from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from oauth2client.service_account import ServiceAccountCredentials


class Project:

    def __init__(self, metadata_gs_id, admin_email, info_file):
        self.project_name = ''
        self.metadata_gs_id = metadata_gs_id
        self.folder = None
        self.admin_email = admin_email
        self.info_file = info_file
        self.client = self.service = None
        self.metadata = {}
        self.accounts = {}
        self.cost = self.inner_cost = None

    def __get_metadata(self):
        metadata_gs = self.client.open_by_key(self.metadata_gs_id)
        for sheet in metadata_gs:
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
        spreadsheet1 = self.client.create(f'Расчет стоимости проект {self.project_name}')
        spreadsheet1.share(
                            value=self.admin_email,
                            perm_type='user',
                            role='owner',
                            notify=False,
                            email_message=None
                            )
        self.cost = spreadsheet1

        spreadsheet2 = self.client.create(f'Внутренние расчеты стоимости проекта {self.project_name} (конфиденциально)')
        spreadsheet2.share(
                            value=self.admin_email,
                            perm_type='user',
                            role='owner',
                            notify=False,
                            email_message=None
                            )
        self.inner_cost = spreadsheet2

    def __create_accounts(self):
        for developer in self.metadata['participants']:
            spreadsheet = self.client.create(f'Учет трудозатрат {developer["Имя"]} проект {self.project_name}')
            spreadsheet.share(
                                value=self.admin_email,
                                perm_type='user',
                                role='owner',
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
            except gspread.exceptions.APIError:
                print(f'There is no Google account associated with this email address: {developer["email"]}')

            worksheet = spreadsheet.add_worksheet('timesheet', 300, 12)
            spreadsheet.del_worksheet(spreadsheet.sheet1)
            cell_list = worksheet.range('A1:D1')
            for cell, value in zip(cell_list, ('Дата', 'Название работ', 
                                   'Объем работ, часов', 'Статус оплаты (заполняется только РП)')):
                cell.value = value
            worksheet.update_cells(cell_list, value_input_option='USER_ENTERED')
            self.accounts[developer['Имя']] = spreadsheet
            time.sleep(3)

    def __create_worksheets(self, spreadsheet):
        for account in self.accounts:
            ws_work = spreadsheet.add_worksheet(f'Работы {account}', 300, 12)
            ws_work.update_cell(1, 1, f'=IMPORTRANGE("{self.accounts[account].id}", "timesheet!A:D")')

    def __update_cost(self):
        spreadsheet = self.cost
        ws_result = spreadsheet.add_worksheet('Итог', 300, 12)
        spreadsheet.del_worksheet(spreadsheet.sheet1)

        self.__create_worksheets(spreadsheet)

        cell_list = ws_result.range('A1:F1')
        for cell, value in zip(cell_list, ('Исполнитель', 'Должность', 'К оплате, часов', 'Ставка, рублей', 
                               'К оплате, рублей', 'Примечания')):
            cell.value = value
        ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')

        row = 2
        for developer in self.metadata['participants']:
            cell_list = ws_result.range(row, 1, row, 6)
            for cell, value in zip(cell_list, (
                    developer['Имя'],
                    developer['Должность'],
                    '''=СУММЕСЛИ('Работы {0}'!$D$2:$D, "Акт", 'Работы {0}'!$C$2:$C)'''.format(developer['Имя']),
                    developer['Ставка внешняя'],
                    '=$C{0}*$D{0}'.format(str(row)),
                    developer['Комментарии к затратам'])):
                cell.value = value
            ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')
            row += 1

        for resource in self.metadata['resources']:
            for col, value in zip((1, 5, 6), (
                    resource['Наименование'],
                    resource['Стоимость, рублей'],
                    resource['Комментарии к затратам'])):
                ws_result.update_cell(row, col, value)
            row += 1
        ws_result.update_cell(row, 1, 'Итого')
        ws_result.update_cell(row, 5, f'=СУММ(E2:E{str(row-1)})')

    def __update_inner_cost(self):
        spreadsheet = self.inner_cost
        ws_result = spreadsheet.add_worksheet('Итог', 300, 26)
        spreadsheet.del_worksheet(spreadsheet.sheet1)

        self.__create_worksheets(spreadsheet)

        cell_list = ws_result.range('A2:A6')
        for cell, value in zip(cell_list, (
                                'К оплате заказчиком по актам, рублей',
                                'К оплате разработчикам по актам, рублей',
                                'Фонд встреч, рублей',
                                'Бонусный фонд, рублей',
                                'Предприятию включая операционные расходы и фонд встреч, рублей')):
            cell.value = value
        ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')

        cell_list = ws_result.range('C2:C6')
        for cell, value in zip(cell_list, ('=$H$12',
                                           '=$I$12',
                                           '',
                                           '=((($C$2*0.9)-$C$3)/2)-$C$4',
                                           '=$C$2-$C$3-$C$5')):
            cell.value = value
        ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')

        for col, value in zip((3, 11, 17, 19), 
                              ('Основной ФОТ', 'Премиальный ФОТ', 'К оплате', 'Общая задолженность')):
            ws_result.update_cell(8, col, value)

        headers = ('Исполнитель', 'Должность', 'Ставка заказчику, рублей', 'Ставка разработчику, рублей',
                   'Ставка для документов заказчику, рублей', 'По актам, часов', 'Часов для документов',
                   'Стоимость для документов, рублей', 'К оплате разработчику по актам, рублей', '',
                   'Вес вклада, баллы', 'Вклад с учетом веса, баллы', r'Предложение доли фонда, %%',
                   r'Фактическая доля фонда, %%', 'Доля фонда, рублей', '',
                   'К оплате разработчику по актам c премией, рублей', '', 'Всего часов к оплате',
                   'Общая задолженность заказчика, рублей', 'Общая задолженность разработчику, рублей',
                   'Общая задолженность предприятию, рублей')

        cell_list = ws_result.range('A9:V9')
        for cell, value in zip(cell_list, headers):
            cell.value = value
        ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')

        row = 10
        sum_row = row + len(self.metadata['participants'])
        for developer in self.metadata['participants']:
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
            cell_list = ws_result.range('A{0}:V{0}'.format(str(row)))
            for cell, value in zip(cell_list, values):
                cell.value = value
            ws_result.update_cells(cell_list, value_input_option='USER_ENTERED')
            row += 1

        ws_result.update_cell(row, 1, 'Итого')
        for col in ('F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O', 'Q', 'S', 'T', 'U', 'V'):
            ws_result.update_acell(f'{col}{str(row)}', '=СУММ({0}10:{0}{1})'.format(col, str(row-1)))

    def __move_to_folder(self):

        def move(file_id):
            self.service.files().update(fileId=file_id,
                                        addParents=self.folder['id'],
                                        fields='id, parents').execute()

        for sheet_id in (self.metadata_gs_id, self.cost.id, self.inner_cost.id):
            move(sheet_id)

        for account in self.accounts.values():
            move(account.id)

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
        self.__create_accounts()
        self.__update_cost()
        time.sleep(100)
        self.__update_inner_cost()
        self.__move_to_folder()


if __name__ == '__main__':

    data = json.load(open('admin_data.json'))

    my_project = Project(data['metadata_gs_id'], data['admin_email'], data['info_file'])
    my_project.main()

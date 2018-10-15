# GSmetadata

Скрипт для создания Гугл таблиц проекта, используя его метаданные
Используется Google Sheets API, Google Drive API, Python 3.7 и библиотеки, указанные в requirements.py

Для начала работы необходимо зайти в свой Гугл аккаунт и перейти по ссылке: https://developers.google.com/drive/api/v3/quickstart/python
Нажимаем "Enable the drive API", создаем новый проект, скачиваем credentials.json и сохраняем его в папку со скриптом. Переходим в консоль API (https://console.cloud.google.com), подключаем еще Google Sheets API.
Далее заходим в учетные данные и создаем сервисный аккаунт, который создает нам еще один json файл, который тоже кидаем в папку со скриптом. Из этого json берем client_email и в нашей Гугл таблице Metadata даем этому email разрешение на чтение.
И самое последнее. В папке скрипта изменяем файл admin_data.json, куда записываем id Гугл таблицы Metadata (копируем из url), свой email Гугл аккаунта и название json файла, который нам дал сервисный аккаунт. Должно получиться что-то такое:
{
    "metadata_gs_id": "veryVERYlongID",
    "admin_email": "admin@gmail.com",
    "info_file": "ProjectName-verylongid.json"
}

Все готово к работе!

К сожалению, у Гугла есть ограничения на запросы в определенный отрезок времени на бесплатный API, а также тормозящая система предоставления доступа. Поэтому примерное время работы скрипта рассчитывается как время = (количество участников в проекте * 5 + 120) секунд. Прилагаю ссылки, чтоб вы не подумали, что это моя прихоть:

https://www.codesd.com/item/error-500-when-using-the-google-drive-api-to-update-permissions.html
https://stackoverflow.com/questions/30027221/getting-500-error-when-using-google-drive-api-to-update-permissions
https://developers.google.com/sheets/api/limits

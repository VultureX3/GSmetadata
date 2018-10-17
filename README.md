# GSmetadata

Скрипт для создания Гугл таблиц проекта, используя его метаданные
Используется Google Sheets API, Google Drive API, Python 3.7 и библиотеки, указанные в requirements.txt

Для начала работы необходимо зайти в свой Гугл аккаунт и перейти по ссылке: https://developers.google.com/drive/api/v3/quickstart/python
Нажимаем "Enable the drive API", создаем новый проект, скачиваем credentials.json и сохраняем его в папку со скриптом. Переходим в консоль API (https://console.cloud.google.com), подключаем еще Google Sheets API.
Далее заходим в "АПИ и сервисы" - "учетные данные" и создаем сервисный аккаунт, который создает нам еще один json файл, который тоже кидаем в папку со скриптом. Из этого json берем client_email и в нашей Гугл таблице Metadata даем этому email разрешение на чтение.
И самое последнее. В папке скрипта изменяем файл admin_data.json, куда записываем id Гугл таблицы Metadata (копируем из url), свой email Гугл аккаунта и название json файла, который нам дал сервисный аккаунт. В четвертый параметр "sharing" записывается 0, если мы не хотим включить автоматическое предоставление доступа всем участникам проекта к таблицам их расчета трудозатрат, и любое другое число, если хотим. Должно получиться что-то такое:

{
	"metadata_gs_id": "veryVERYlongID",
    "admin_email": "admin@gmail.com",
    "info_file": "ProjectName-verylongid.json",
    "sharing": 0
}

Все готово к работе!

Чтобы запустить скрипт из командной строки, просто заходим в папку проекта, где лежит скрипт, и пишем "python3 script.py".

К сожалению, у Гугла есть ограничения на запросы в определенный отрезок времени на бесплатный API, а также тормозящая система предоставления доступа. Поэтому на каждого участника проекта уходит 10-15 секунд. Это время увеличивается в 1,5 - 2 раза, если поставить параметр "sharing" отличным от нуля. Прилагаю ссылки по теме:

https://www.codesd.com/item/error-500-when-using-the-google-drive-api-to-update-permissions.html
https://stackoverflow.com/questions/30027221/getting-500-error-when-using-google-drive-api-to-update-permissions
https://developers.google.com/sheets/api/limits

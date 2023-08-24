# Скрипт импорта информации

## Описание 
Данный скрипт предназначен для автоматизации загрузки данных сотрудников в SharePoint. Скрипт загружает данные из текстового файла, содержащего информацию о сотрудниках, и добавляет их в соответствующие списки SharePoint. Скрипт загружает данные в следующие списки SharePoint: "Организационно-штатная структура", "Наименование должностей","Физические лица". Для каждой загруженной записи создается соответствующий элемент в списке, а также производится связывание элементов списков с помощью полей связи (lookup fields). При загрузке происходит проверка на уже имеющиеся данные в списке, чтобы избежать дубликатов.

## Требования
Скрипт использует текстовый файл для загрузки информации о сотрудниках. Файл должен содержать следующие поля:
 - Surname
 - FirstName
 - SecondName
 - Fullname
 - Subdivision
 - Staffname
 - Email
 - Login

Поля разделены табуляцией ('\t').

## Использование
1. Запустите скрипт в среде PowerShell.

2. Скрипт подключается к указанному веб-сайту SharePoint с использованием текущих учетных данных и выполняет следующие действия:
    - Определяет все уникальные подразделения и должности.
    - Загружает подразделения и добавляет их как группы Sharepoint, с проверкой на уже существующие.
    - Загружает должности, с проверкой на уже существующие.
    - Загружает физические лица, с проверкой на уже существующие.
    - Загружает сотрудников в ОШС по уже имеющимся ID физических лиц.
    - Добавляет пользователей в созданные группы пользователей Sharepoint.
    - Обновляет подразделения ОШС, проставляет поле Руководитель и Группа.
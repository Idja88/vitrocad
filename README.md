# Автоматизация SharePoint и VitroCAD с помощью PowerShell PnP

Репозиторий содержит набор скриптов PowerShell для автоматизации различных бизнес-задач и процессов, связанных с SharePoint и системой VitroCAD, с использованием модуля SharePoint PnP. Скрипты используют столбцы и элементы системы, идущие в типовой конфигурации системы VitroCAD.

## Введение

SharePoint PnP (Patterns and Practices) - это сборник библиотек, инструментов и обучающих материалов, которые позволяют работать с SharePoint более эффективно и удобно. Модуль SharePoint PnP PowerShell облегчает выполнение операций с SharePoint, таких как создание и удаление сайтов, управление списками и библиотеками документов, а также другие общие задачи.

VitroCAD - это система управления инженерными данными и документацией (EDMS), основанная на платформе Microsoft SharePoint. VitroCAD обеспечивает управление жизненным циклом документов, совместную работу, автоматизацию бизнес-процессов и интеграцию с другими системами.

## Структура репозитория

В репозитории представлены следующие скрипты:

1. **OrgStructureSync**
2. **OrgStructureVacation**
3. **IssuesMailSender**
4. **AddUniqueRoles**
5. **ImportFromCsv**

## Требования

- PowerShell 5.1 или выше
- Microsoft.SharePoint.PowerShell модуль: `Add-PSSnapin Microsoft.SharePoint.PowerShell`
- PnP PowerShell модуль: `Install-Module -Name SharePointPnPPowerShell2013`

## Использование

1. Клонируйте репозиторий или скачайте архив с файлами
2. Откройте скрипт из соответствующей папки, который соответствует вашей задаче, и настройте переменные в начале файла
3. Запустите скрипт в PowerShell с соответствующими правами доступа

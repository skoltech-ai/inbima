# inbima

> This is the **IN**telligent system for **BI**bliographic data **MA**nagement.


## !!! Note

This is a draft version of a software product that demonstrates the idea and basic functionality. The code is dirty! The capabilities of automatic report generation are currently minimal (only for proof of concept). The information contained in the [excel database](https://docs.google.com/spreadsheets/d/1jz76t2bRMzlNqL315SUf1WKr45lWu-c_/edit?usp=sharing&ouid=102021586566196668105&rtpof=true&sd=true) is collected automatically, and in the future it will be checked manually. Support for parsing scientific conferences will also be added. The package text description below is also a draft and may contain typos and inaccuracies, and it maybe has bad language style.


## How to use

1. Explore the information about team members and publications presented in the [excel database](https://docs.google.com/spreadsheets/d/1jz76t2bRMzlNqL315SUf1WKr45lWu-c_/edit?usp=sharing&ouid=102021586566196668105&rtpof=true&sd=true). This database represents a single source of truth for the team members, their publications and conferences. It will be automatically downloaded from the Google Drive when the python script from this repository is running.
    > In the future, if necessary, of course, the excel file can be replaced with a relational database. However, at the moment, using the excel file seems to be optimal, since the data and view are combined (no need to develop a system for visualization).
2. Download this [repository](https://github.com/SkoltechAI/inbima).
3. Open downloaded folder `inbima` in terminal (console).
4. Install dependencies by the command (`python 3.6+` is required)
    ```bash
    pip install -r requirements.txt
    ```
5. Run the main script by the command
    ```bash
    python inbima.py
    ```
6. While the script work the [excel database](https://docs.google.com/spreadsheets/d/1jz76t2bRMzlNqL315SUf1WKr45lWu-c_/edit?usp=sharing&ouid=102021586566196668105&rtpof=true&sd=true) will be downloaded from Google Drive into temporary folder `export_DATETIME`. All generated reports will be also saved into this folder.
7. The folder `export_DATETIME` will contain automatically generated reports using the default settings of the `inbima`.
    > In the current draft version of the program, only a bibliographic report for one member of the team is generated, as well as a report for articles within one grant. In the future, the type of the generated report can be specified by setting the appropriate flags in the excel file on the `INBIMA` sheet.
8. [TODO (reading options from excel file is not supported yet)] To customize filters and other program settings, open the excel database file `cait.xlsx` from the same folder `export_DATETIME` and modify the `INBIMA` list in the file.
    > At the moment this is not relevant, since flags from excel file are not used in the current version of the code.
9. Re-run the script by adding the flag `--last` (or `-l`)
    ```bash
    python inbima.py -l
    ```
    > When launched with this flag, the same folder `export_DATETIME` will be used (the folder with the maximum date value will be selected). When the script is running, the data and settings specified in the database `cait.xlsx` from this folder will be used. Note that all reports from the previous run will be removed from the folder! In order to avoid deleting previously generated reports, you can create a new folder `export_DATETIME` (with the highest `DATETIME` value) and copy the excel database `cait.xlsx` into it.
10. [TODO (reading options from excel file is not supported yet)] The folder `export_DATETIME` will contain automatically generated reports using the customized settings of the `inbima`.
    > At the moment this is not relevant, since flags from excel file are not used in the current version of the code.


## Content

- `inbima.py` - the main script, that contains `InBiMa` class;
- `parser_wos.py` - this script was used for automatic parsing of the `.bib` files manually uploaded from WoS. If you have any questions regarding the usage of this script, please contact the author;
- `bib_wos` - this folder contains the `.bib` files for the core team members manually uploaded from WoS;
- `scimagojr 2020.csv` - this file contains information about scientific journals. It was downloaded from [Scimago Journal & Country Rank](https://www.scimagojr.com/journalrank.php), and it was used for automatic parsing of the papers;
- `requirements.txt` - this file contains project dependencies (python libraries);
- `LICENSE.txt` - this is a formal license with fairly broad rights. However, if you want to use this system, please contact the authors first and get their explicit consent.


## Authors

- [Andrei Chertkov](https://github.com/AndreiChertkov) (a.chertkov@skoltech.ru).

> We will be glad to all volunteers ;)


## Instructions for adding and updating publications

> At the moment, the instruction has been prepared only in Russian.

> Уже проверенные статьи/журналы/конференции можно использовать в качестве образца (для таких статей в столбце `Checked` указана дата проверки и автор проверки).

### Добавление/уточнение публикации (лист `papers`)

> См. примеры обработанных первых 5 статей в списке.

1. Открыть папку [inbima](https://drive.google.com/drive/folders/1GK1eDkU0vqLz8gFMXxJuwqqCzSjmfhZx?usp=sharing) на Google Drive, содержащую базу данных в excel формате (далее БД), а также ряд вложенных папок со вспомогательной информацией (тексты публикаций, скриншоты подтверждений, и т.д.).
2. Открыть БД [cait.xlsx](https://docs.google.com/spreadsheets/d/1jz76t2bRMzlNqL315SUf1WKr45lWu-c_/edit?usp=sharing&ouid=102021586566196668105&rtpof=true&sd=true). Данные, содержащиеся в этой базе, представляют "единый консистентный источник истины" по членам коллектива, их публикациям и конференциям. Вносить изменения в БД стоит с осторожностью!
    > Редактирование документа по ссылке недоступно. На данный момент редактировать БД может лишь автор.
3. Перейти на лист `papers` в БД (далее работаем на этом листе по умолчанию). Выбрать публикацию для уточнения, либо собрать первичную информацию по новой добавляемой публикации и создать новую строку в начале листа.
4. Подготовить уникальный идентификатор публикации по стандартному правилу bibtex именования `авторГОДпервое_слово_названия_статьи` (всё в нижнем регистре; "предлоги" в начале названия пропускаются; дефис в первом слове удаляется при наличии и части объединяются в одно слово), **проверить отсутствие совпадений** (при наличии - добавить второе слово из названия статьи), задать его в поле `ID` вместо автоматически сгенерированного идентификатора.
5. Найти страницу публикации на сайте издательства, задать url-адрес в поле `Site`.
6. Сделать скриншот страницы на сайте издательства (в формате `png`), в качестве имени файла использовать `ID`, заданное выше. Сохранить скриншот в папку [papers_screen](https://drive.google.com/drive/folders/1MkuFZrBCRv4EIkJzIZVld5EE-krznFVw?usp=sharing) на Google Drive. Используя кнопку `Share` на Google Drive, скопировать ссылку на загруженный файл и указать ее в поле `Screen`.
7. Найти pdf-файл публикации (в формате `pdf`), в качестве имени файла использовать `ID`, заданное выше. Сохранить файл в папку [papers_text](https://drive.google.com/drive/folders/1jNrhKlyacDc07wOyHPhkL2UziQLhmHZd?usp=sharing) на Google Drive. Используя кнопку `Share` на Google Drive, скопировать ссылку на загруженный файл и указать ее в поле `Text`.
    - Если найти файл не удалось, то в поле `Note` указать `TEXT-NO`.
    - Если удалось найти только текст с архива, то в поле `Note` указать `TEXT-ARXIV`.
8. Проверить/задать поля `Title`, `Year`, `Volume`, `Number`, `Pages`, `DOI`.
9. Проверить/задать поле `Authors`. Авторы указываются в формате `Sun H., Jin J., Xu R., Cichocki A.`. Проверить/задать поле `Authors Parsed` (копируется содержимое поля `Authors`, затем члены коллектива заменяются на идентификатор, указанный на листе `team`, например, `Sun H., Jin J., Xu R., #cichocki`).
10. Найти журнал публикации на листе `journals` в БД и задать соответствующее название в поле `Journal`. Если журнала в списке нет или журнал ранее не проверялся (пустое поле `Checked`), то его необходимо прежде добавить/проверить, используя инструкцию по дабавлению журнала (см. ниже).
11. Поля `Grant` и `Grant_str` пока не заполняются (в них содержатся автоматически сгенерированные/распарсенные значения). /* TODO */
12. Указать в поле `Checked` на листе `papers` дату проверки и автора проверки в формате `2021-10-29-ac`.

### Добавление/уточнение журнала (лист `journals`)

> См. примеры обработанного первого журнала в списке, а также "International Journal of Neural Systems".

1. Открыть папку [inbima](https://drive.google.com/drive/folders/1GK1eDkU0vqLz8gFMXxJuwqqCzSjmfhZx?usp=sharing) на Google Drive и БД [cait.xlsx](https://docs.google.com/spreadsheets/d/1jz76t2bRMzlNqL315SUf1WKr45lWu-c_/edit?usp=sharing&ouid=102021586566196668105&rtpof=true&sd=true).
2. Перейти на лист `journals` в БД (далее работаем на этом листе по умолчанию). Выбрать журнал для уточнения, либо собрать первичную информацию по новому журналу и создать новую строку в начале листа.
3. Найти сайт журнала, скопировать адрес в поле `Site`, а корректное название журнала (в правильном регистре) в поле `Title`.
3. Открыть [сайт WoS с информацией по публикациям](https://apps.webofknowledge.com). Организовать поиск по любой статье из данного журнала (можно использовать статью из данного журнала ранее добавленную на лист `papers` БД). На странице с результатами поиска кликнуть на название журнала. После отображения краткой информации по журналу сделать скриншот всей страницы (в формате `png`), в качестве имени файла использовать имя журнала, заданное выше. Сохранить скриншот в папку [journals_screen_wos](https://drive.google.com/drive/folders/1HH9WS_lWdA-xTAaHmHuu0b_8VICbqUbT?usp=sharing) на Google Drive. Используя кнопку `Share` на Google Drive, скопировать ссылку на загруженный файл и указать ее в поле `Screen WoS`. Скопировать (это нетривиально) импакт-фактор журнала в поле `WoS Impact`, а названия категорий в соответствующее поля квартиля (`WoS Q1`, `WoS Q2`, `WoS Q3` или `WoS Q4`).
    > Возможен поиск журналов на специализированном [сайте WoS с информацией по журналам](https://jcr.clarivate.com), однако там на персональной страничке журнала содержится слишком много информации, и делать скриншот очень неудобно, поэтому вероятно лучше использовать предлагаемый здесь способ.

- TODO Продумать сохранение информации о журналах из базы Scopus.
- TODO Переименовать поля типа `WoS Q1` на более удачные (как минимум, лучше `Q1 WoS`, `Q1 SJR`, etc.).
- TODO Уточнить, как корректно добавлять категории журнала, по которому расчитывается импакт фактор.

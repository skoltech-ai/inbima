# inbima

> This is the **IN**telligent system for **BI**bliographic data **MA**nagement. Note that this is a draft version of a software product that demonstrates the idea and basic functionality (the code is dirty and the capabilities of automatic report generation are currently minimal). If you want to use this system, please contact the author first and get their explicit consent.


## How does it work

1. We have downloaded the complete bibliographic information on all key members of the research team of our center from [Scopus](https://www.scopus.com/authid/detail.uri?authorId=8529104000) and [WoS](https://www.webofscience.com/wos/author/record/565227,36803309,44328120) databases as well as information on scientific journals for the current year from [Scopus journals list](https://www.scopus.com/sources), [WoS journals list]() and [SJR](https://www.scimagojr.com/journalrank.php).
2. We automatically downloaded all collected publications, linked them with team members, journal ratings and our grants (**TODO** the list of grants has not yet been compiled in full), and saved them in our own online [excel-database](https://docs.google.com/spreadsheets/d/17Yi1Jg6DF3k7pFWm-N9jYhcU4tEIZCcv9clq3y0JHVc/edit?usp=sharing). The date of downloading information from scientific databases and uploading to our own excel-database is November 1, 2021, after which all new publications will be added to the excel-database manually (with appropriate careful verification).
3. [**TODO**] We manually checked all publications and journals which were automatically added into the excel-database and supplemented them with the necessary reference information (link to the publisher's website, publication text, etc.).
    > Note that for any publication, a single manual check is sufficient, after that the corresponding row in the database sheet is protected from changes (in the future, only the team members in the list of the authors may be updated, if required), but for each journal it is necessary (at least) once a year to update its rating.
4. [**In progress**] We have developed a program code to download publications from our excel-database and to automatically generate various reports in various formats (`MS Word`, `MS Excel`, `bibtex`, `json`, etc.) with powerfull filters. This auto generated reports can potentially be used for preparing commercial proposals, grant applications, reporting on grants, for presenting information on websites and for preparing various statistics in tabular and graphical forms.
5. [**TODO**] We have manually prepared a list of talks at conferences and conference proceedings made by the team members. This information can also be used to generate reports.
6. [**TODO**] We have automatically uploaded all the arxiv preprints of the research team members. This information can also be used to generate reports (after manual check?).
7. [**TODO**] We have organized an interactive form for collecting information on new publications, talks at conferences and other achievements of the research team members. The relevant information is then manually checked, supplemented and added to our excel-database.


## How to use

1. Download this [repository](https://github.com/SkoltechAI/inbima).
2. Open downloaded folder `inbima` in terminal (console) and install dependencies by the command (`python 3.6+` is required) `pip install -r requirements.txt`.
3. Run the main script by the command `python inbima.py -f`
    > With the flag `-f` only working folder `./export/export_CURRENT_TIMESTAMP` will be created (of course this operation can be performed manually without running the script). During the further operation of the program, **the folder with the maximum timestamp will be selected as a working folder by the script for reading the database and writing reports**.
4. Download the [excel-database](https://docs.google.com/spreadsheets/d/17Yi1Jg6DF3k7pFWm-N9jYhcU4tEIZCcv9clq3y0JHVc/edit?usp=sharing) manually and save it in the working folder (`./export/export_LAST_TIMESTAMP`) as `cait.xlsx`.
    > Click in Google Drive menu `File > Download > Microsoft Excel (.xlsx)` to download the excel file.
5. [**TODO**] To customize the kinds/filters for reports, open the downloaded excel-database file `cait.xlsx` and modify the options on the `INBIMA` sheet.
    > In the current version of the program these kinds/filters are not used, and only a bibliographic reports for the all team members are generated, as well as a report for publications within one grant.
6. Run the main script by the command `python inbima.py`.
7. The folder `./export/export_LAST_TIMESTAMP` will contain all automatically generated reports.
8. There is one more usefull option. If you run the script with the flag `-j` as `python inbima.py -j 'JOURNAL_NAME_IN_QUOTES'` then the full info about specified journal will be logged to console.
    > If an incorrect journal title is entered, then the titles of the 10 most similar titles will be displayed in the console (then you can use the correct title and rerun the script with it). Note that at the moment this flag `-j` is experimental, and in the future the more detailed information will be presented.


## Author

- [Andrei Chertkov](https://github.com/AndreiChertkov) (a.chertkov@skoltech.ru).


## Contributors

- ... we will be glad to all volunteers ;)

# inbima

> This is the **IN**telligent system for **BI**bliographic data **MA**nagement.


## !!! Note

- This is a draft version of a software product that demonstrates the idea and basic functionality.
- The code is dirty!
- The capabilities of automatic report generation are currently minimal (only for proof of concept).
- The script `parser_wos.py` fills the database automatically using records from ишиеуч files. At the moment, the code in this file is not up-to-date and the script may not work correctly.
- Support for parsing scientific conferences will be added later.
- If you want to use this system, please contact the authors first and get their explicit consent.


## How to use

1. Download this [repository](https://github.com/SkoltechAI/inbima).
2. Open downloaded folder `inbima` in terminal (console) and install dependencies by the command (`python 3.6+` is required)
    ```bash
    pip install -r requirements.txt
    ```
3. Run the main script by the command
    ```bash
    python inbima.py -f
    ```
    > With the flag `-f` only working folder `./export/export_CURRNET_TIMESTAMP` will be created (this operation can be performed manually without running the script). During the further operation of the program, the folder with the maximum timestamp will be selected as a working folder.
4. Download the [excel database](https://docs.google.com/spreadsheets/d/17Yi1Jg6DF3k7pFWm-N9jYhcU4tEIZCcv9clq3y0JHVc/edit?usp=sharing) manually and save it in the working folder as `cait.xlsx`.
    > Click in Google Drive menu `File > Download > Microsoft Excel (.xlsx)` to download the excel file.
5. Run the main script by the command
    ```bash
    python inbima.py
    ```
6. The folder `./export/export_LAST_TIMESTAMP` will contain automatically generated reports using the default settings of the `inbima`.
    > In the current draft version of the program, only a bibliographic reports for the all team members are generated, as well as a report for articles within one grant (`megagrant1`). In the future, it will be possible to specify the type of the generated report by setting the appropriate parameters and flags in the excel file on the `INBIMA` sheet.


## Authors

- [Andrei Chertkov](https://github.com/AndreiChertkov) (a.chertkov@skoltech.ru).

> We will be glad to all volunteers ;)

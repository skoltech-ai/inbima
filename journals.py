import csv
from Levenshtein import distance
import openpyxl


from utils import load_sheet
from utils import log


REF_SCO_PATH = './journals/CiteScore.xlsx'
REF_SCO_SHEET = 'DATA'
REF_SCO_INDEX_NAME = 'CiteScore 2020'
REF_SJR_PATH = './journals/scimagojr 2020.csv'


class Journals():
    def __init__(self, data={}):
        self.data = data

    def find(self, title, journals=None):
        journals = journals or self.data

        res = []
        for journal in journals.values():
            dist = distance((title or '').lower(), journal['title'].lower())
            res.append([journal, dist])
        res.sort(key=lambda x: x[1])

        if len(res) == 0:
            return None, []

        if len(res) == 1:
            if res[0][1] == 0:
                return res[0][0], []
            else:
                return None, [res[0]['title']]

        if res[0][1] == 0 and res[1][1] != 0:
            return res[0][0], []

        titles = [r[0]['title'] for r in res[:10]]
        return None, titles

    def get_journal(self, title=None, issn=None, dist_max=0, dist_max_wrn=1):
        if not issn and (not title or len(title) < 2):
            return

        for title_real, item in self.journals_ref.items():
            t = title_real.lower()
            if issn and issn == item['issn']:
                if title and distance(title.lower(), t) > dist_max:
                    text = 'Journal found by ISSN but titles are different: '
                    text += f'"{title}" is replaced by "{title_real}"'
                    log(text, 'wrn')
                return item
            dist = distance(title.lower(), t)
            if dist > dist_max:
                continue
            if dist >= dist_max_wrn:
                text = f'Journal "{title}" is replaced by "{title_real}"'
                log(text, 'wrn')
            return item

    def load_ref(self):
        self.load_ref_sjr()
        # self.load_ref_sco()
        # for j in self.ref_sjr.values():
        #    self.print_ref_sjr(j)
        # print(self.ref_sjr_quartiles)

    def load_ref_sco(self):
        """Load Scopus journals info from excel file.

        Note:
            See for data https://www.scopus.com/sources.

            We are interested in the following fields of the excel file:
                Col. 01 Title -> 'title' (it will be the key for dict);
                Col. 02 CiteScore 2020 -> 'sco_index'
        """
        self.ref_sco = {}

        wb = openpyxl.load_workbook(REF_SCO_PATH)
        data = load_sheet(wb[REF_SCO_SHEET])
        for row in data.values():
            self.ref_sco[row['title']] = {
                'title': row['title'],
                'sco_index': row[REF_SCO_INDEX_NAME],
                'sco_sjr_index': row['sjr'],
                'sco_area': row['Scopus Sub-Subject Area'],
            }

    def load_ref_sjr(self):
        """Load Scimago Journal Rank from the csv file.

        Note:
            See for data https://www.scimagojr.com/journalrank.php.

            See for info https://www.scimagojr.com/help.php.

            We are interested in the following fields of the csv file:
                No. 00 Rank       -> 'sjr_rank';
                No. 02 Title      -> 'title' (it will be the key for dict);
                No. 03 Type       -> 'sjr_type'
                No. 04 Issn       -> 'issn' (only print ISSN saved);
                No. 05 SJR        -> 'sjr_index';
                No. 06 SJR Best Quartile
                No. 07 H index    -> 'sjr_h_index';
                No. 15 Country    -> 'country';
                No. 17 Publisher  -> 'publisher';
                No. 19 Categories -> 'sjr_q_raw','sjr_q1',...,'sjr_q4','sjr_q0'

        """
        self.ref_sjr = {}
        self.ref_sjr_quartiles = []

        def parse_issn(issn):
            if not issn or len(issn) < 8:
                return ''
            return issn[-8:-4] + '-' + issn[-4:]

        def parse_quartiles(data, kind):
            res = []
            for item in data.split('; '):
                if item.endswith('(Q1)'):
                    if kind == 'Q1': res.append(item[:-5])
                elif item.endswith('(Q2)'):
                    if kind == 'Q2': res.append(item[:-5])
                elif item.endswith('(Q3)'):
                    if kind == 'Q3': res.append(item[:-5])
                elif item.endswith('(Q4)'):
                    if kind == 'Q4': res.append(item[:-5])
                else:
                    if kind == 'Q0': res.append(item)

            for name in res:
                if not name in self.ref_sjr_quartiles:
                    self.ref_sjr_quartiles.append(name)

            return res

        with open(REF_SJR_PATH, newline='') as f:
            for row in [r for r in csv.reader(f, delimiter=';')][1:]:
                self.ref_sjr[row[2]] = {
                    'sjr_rank': row[0],
                    'title': row[2],
                    'sjr_type': row[3],
                    'issn': parse_issn(row[4]),
                    'sjr_index': row[5] or 0,
                    'sjr_h_index': row[7],
                    'country': row[15],
                    'publisher': row[17],
                    'sjr_q_raw': row[19],
                    'sjr_q1': parse_quartiles(row[19], 'Q1'),
                    'sjr_q2': parse_quartiles(row[19], 'Q2'),
                    'sjr_q3': parse_quartiles(row[19], 'Q3'),
                    'sjr_q4': parse_quartiles(row[19], 'Q4'),
                    'sjr_q0': parse_quartiles(row[19], 'Q0'),
                }

    def log_ref(self, title):
        journal, titles = self.find(title, self.ref_sjr)

        if journal:
            self.log_ref_sjr(journal)
        else:
            text = '\n' + '=' * 60 + ' ' + '\n'
            text += 'I can not find this journal. '
            if len(titles):
                text += 'The most similar are:\n\n'
                for title in titles:
                    text += f'>>> "{title}"\n'
            text += '-' * 60 + '\n'

            print(text)

    def log_ref_sjr(self, journal):
        text = '\n' + '=' * 60 + ' '
        text += 'SJR journal info' + '\n'

        v = journal['issn'] or ' '*9
        text += f'[{v}] '

        v = journal['title']
        text += f'|{v}|\n'

        v = ''
        v += (journal['country'] or '') + '; '
        v += (journal['publisher'] or '') + '; '
        v += (journal['sjr_type'] or '')
        text += f'            |{v}|\n'

        v = journal['sjr_rank']
        text += f'>>> Rank    : {v}\n'

        v = journal['sjr_index']
        text += f'>>> Index   : {v}\n'

        v = journal['sjr_h_index']
        text += f'>>> H-Index : {v}\n'

        text += '>>> Q1      : ' + ('; '.join(journal['sjr_q1'])) + '\n'
        text += '>>> Q2      : ' + ('; '.join(journal['sjr_q2'])) + '\n'
        text += '>>> Q3      : ' + ('; '.join(journal['sjr_q3'])) + '\n'
        text += '>>> Q4      : ' + ('; '.join(journal['sjr_q4'])) + '\n'
        text += '>>> Q0      : ' + ('; '.join(journal['sjr_q0'])) + '\n'

        text += '-' * 60 + '\n'

        print(text)

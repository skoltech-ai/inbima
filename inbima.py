import matplotlib.pyplot as plt
import openpyxl
import sys


from fs import FS
from journals import Journals
from utils import load_sheet
from utils import log
from word import Word


YEARS = [2017, 2018, 2019, 2020, 2021]


class InBiMa():
    def __init__(self, is_new_folder=False):
        self.fs = FS(is_new_folder)
        if is_new_folder: return

        self.wb = openpyxl.load_workbook(self.fs.get_path('cait.xlsx'))
        log('Excel file is opened', 'res')

        self.team = load_sheet(self.wb['team'])
        self.grants = load_sheet(self.wb['grants'])
        self.papers = load_sheet(self.wb['papers'])
        self.journals = Journals(load_sheet(self.wb['journals']))
        # self.journals_ref = self.load_journals_ref()
        self.task = {
            'authors': ['#cichocki'],
            'grants': ['#megagrant1'],
        }
        log('Excel file is parsed', 'res')

        # log('Journal info is loaded', 'res')

        for uid in self.team.keys():
            self.task['authors'] = [uid]
            self.export_word_cv()
        self.export_grant_papers()
        self.export_stat()

    def export_word_cv(self):
        if len(self.task.get('authors', [])) != 1:
            text = 'export_word_cv (task should contain only one author)'
            log(text, 'err')
            return

        person = self.team.get(self.task['authors'][0])
        if person is None:
            text = 'export_word_cv (invalid team member uid in task)'
            log(text, 'err')
            return

        uid = person['id']
        stat = self.get_papers_stat(uid, YEARS)
        photo_logo = self.fs.download_photo_logo()
        photo_person = self.fs.download_photo(uid[1:], person.get('photo'))

        self.word = Word(YEARS, self.get_papers)
        self.word.add_person_info(person, photo_person, photo_logo)
        self.word.add_person_stat(stat)
        self.word.add_note(is_grant=True)
        self.word.add_break()
        self.word.add_paper_list(stat, author=person['id'])

        fname = 'CAIT_' + person['surname'] + '_' + person['name'] + '.docx'
        fpath = self.fs.get_path(fname)
        self.word.save(fpath)
        log(f'Document "{fpath}" is saved', 'res')

    def export_grant_papers(self):
        if len(self.task.get('grants', [])) != 1:
            text = 'export_grant_papers (task should contain only one grant)'
            log(text, 'err')
            return

        grant = self.grants.get(self.task['grants'][0])
        if grant is None:
            text = 'export_grant_papers (invalid grant uid in task)'
            log(text, 'err')
            return

        uid = grant['id']
        stat = self.get_papers_stat(years=YEARS, grant=uid)
        photo_logo = self.fs.download_photo_logo()

        head = grant.get('head', '')
        head = self.team[head]

        self.word = Word(YEARS, self.get_papers)
        self.word.add_grant_info(grant, head, photo_logo)
        self.word.add_note(is_grant=True)
        self.word.add_break()
        self.word.add_paper_list(stat, grant=uid, with_links=True)

        fname = 'CAIT_' + uid[1:] + '.docx'
        fpath = self.fs.get_path(fname)
        self.word.save(fpath)
        log(f'Document "{fpath}" is saved', 'res')

    def export_stat(self):
        stats = {}
        for uid in self.team.keys():
            if self.team[uid].get('active') != 'Yes':
                continue
            if self.team[uid].get('lead') != 'Yes':
                continue
            stats[uid] = self.get_papers_stat(uid, YEARS)

        for uid, stat in stats.items():
            x = YEARS
            y = [stat[y]['total'] for y in YEARS]
            plt.plot(x, y, marker='o', label=uid)

        plt.legend(loc='best')

        fpath = self.fs.get_path('plot.png')
        plt.savefig(fpath)
        log(f'Figure "{fpath}" is saved', 'res')

    def get_papers(self, author=None, year=None, q=None, grant=None):
        res = {}

        for title, paper in self.papers.items():
            if year and int(year) != int(paper['year']):
                continue

            if author and not author in paper['authors_parsed']:
                continue

            if grant and not grant in paper.get('grant', ''):
                continue

            if q is not None:
                journal = self.journals.data[paper['journal']]
                q1 = journal.get('sjr_q1', '')
                q2 = journal.get('sjr_q2', '')
                if q == 1 and len(q1) < 2:
                    continue
                if q == 2 and (len(q1) >= 2 or len(q2) < 2):
                    continue
                if q == 0 and (len(q1) >= 2 or len(q2) >= 2):
                    continue

            res[title] = paper
            res[title]['journal_object'] = self.journals.data[paper['journal']]

        return res

    def get_papers_stat(self, author=None, years=[], grant=None):
        res = {}

        for year in years:
            res[year] = {
                'q1': len(self.get_papers(author, year, q=1, grant=grant)),
                'q2': len(self.get_papers(author, year, q=2, grant=grant)),
                'q0': len(self.get_papers(author, year, q=0, grant=grant)),
                'total': len(self.get_papers(author, year, grant=grant))
            }
        res['total'] = {
            'q1': sum(res[year]['q1'] for year in years),
            'q2': sum(res[year]['q2'] for year in years),
            'q0': sum(res[year]['q0'] for year in years),
            'total': sum(res[year]['total'] for year in years),
        }
        return res


if __name__ == '__main__':
    args = sys.argv[1:]
    if len(args) == 0:
        ibm = InBiMa()
    elif len(args) == 1 and args[0] == '-f':
        ibm = InBiMa(is_new_folder=True)
    elif len(args) == 2 and args[0] == '-j':
        journals = Journals()
        journals.load_ref()
        journals.log_ref(title=args[1])
    else:
        raise ValueError('Invalid arguments for script')

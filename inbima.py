import csv
from datetime import datetime
import docx
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from docx.shared import Mm, Cm, Pt
import openpyxl
import os
from Levenshtein import distance
import requests
import shutil
import sys


DEBUG = False
JOURNALS_REF_SJR_FILE_PATH = './scimagojr 2020.csv'
YEARS = [2017, 2018, 2019, 2020, 2021]


class InBiMa():
    def __init__(self, use_last=False):
        self.init(use_last)

        self.team = self.load_table('team')
        self.grants = self.load_table('grants')
        self.papers = self.load_table('papers')
        self.journals = self.load_table('journals')
        self.journals_ref = self.load_journals_ref()
        self.task = {
            'authors': ['#cichocki'],
            'grants': ['#megagrant1'],
        }
        self.log('Excel file is parsed', 'res')

        self.export_word_cv()
        self.export_grant_papers()

    def compose_paper_nums(self, paper, end='.'):
        volume = paper.get('volume')
        number = paper.get('number')
        pages = paper.get('pages')
        text = ''
        if volume:
            text += f'{volume}({number})' if number else f'{volume}'
            text += f': {pages}' if pages else ''
            text += end
        elif pages:
            text += f'{pages}'
            text += end
        return text

    def docx_create(self):
        document = docx.Document()

        style_normal = document.styles['Normal']
        style_normal.font.name = 'Cambria'
        style_normal.font.size = Pt(12)

        style_title_name = document.styles.add_style('TitleName',
            WD_STYLE_TYPE.PARAGRAPH)
        style_title_name.base_style = style_normal
        style_title_name.font.name = 'Bookman Old Style'
        style_title_name.font.size = Pt(20)

        table_stat_title = document.styles.add_style('TableStatTitle',
            WD_STYLE_TYPE.PARAGRAPH)
        table_stat_title.base_style = style_normal
        table_stat_title.font.name = 'Bookman Old Style'
        table_stat_title.font.size = Pt(12)

        styles = {
            'normal': style_normal,
            'title_name': style_title_name,
            'table_stat_title': table_stat_title,
        }

        return document, styles

    def docx_build_paper_list(self, document, stat, author=None, grant=None, with_links=False):
        for q in [1, 2, 0]:
            if stat['total'][f'q{q}'] == 0:
                continue
            if q == 1:
                text = 'Publications in Q1-rated journals'
            elif q == 2:
                text = 'Publications in Q2-rated journals'
            else:
                text = 'Other publications'
            document.add_heading(text, level=3)

            for year in YEARS[::-1]:
                if stat[year][f'q{q}'] == 0:
                    continue

                papers = self.get_papers(author, year, q, grant)

                document.add_heading('Year ' + str(year), level=5)

                for paper in papers.values():
                    journal = self.journals[paper['journal']]

                    p = document.add_paragraph(style='List Bullet') # Number
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                    authors = paper['authors'].split(', ')
                    authors_parsed = paper['authors_parsed'].split(', ')

                    for i in range(len(authors)):
                        text = p.add_run(authors[i])
                        if authors_parsed[i] in (author or ''):
                            text.bold = True
                        if not author and authors_parsed[i][0] == '#':
                            text.bold = True
                        if i < len(authors)-1:
                            p.add_run(', ')
                        else:
                            p.add_run('. ')

                    p.add_run(paper['year'] + '. ')
                    p.add_run(paper['title'] + '. ')
                    p.add_run(paper['journal'] + '. ').italic = True
                    p.add_run(self.compose_paper_nums(paper))
                    if with_links:
                        v = paper.get('screen')
                        if v and len(v) > 2:
                            p.add_run(' // ')
                            add_hyperlink(p, 'Screenshot Publisher', v)
                        v = journal.get('screen_wos')
                        if v and len(v) > 2:
                            p.add_run(' // ')
                            add_hyperlink(p, 'Screenshot WoS', v)

            p.add_run('\n\n')

    def docx_build_grant_info(self, grant, document, styles):
        table = document.add_table(rows=1, cols=2)
        table.rows[0].cells[0].width = Cm(6.5)
        table.rows[0].cells[0].height = Cm(20)
        table.rows[0].cells[1].width = Cm(12)
        table.rows[0].cells[1].height = Cm(20)

        table.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = table.rows[0].cells[0].paragraphs[0]
        file_path = self.download_photo_logo()
        if file_path:
            p.add_run().add_picture(file_path, width=Cm(6))

        table = table.rows[0].cells[1].add_table(rows=5, cols=2)
        table.rows[0].cells[0].width = Cm(3)
        table.rows[0].cells[1].width = Cm(9)

        table.rows[0].cells[0].text = 'Grant id'
        table.rows[0].cells[1].text = str(grant.get('id', ''))

        table.rows[1].cells[0].text = 'Head of the project'
        head = grant.get('head', '')
        head = self.team[head]
        head = head['name'] + ' ' + head['surname']
        table.rows[1].cells[1].text = head

        table.rows[2].cells[0].text = 'Grant Number'
        table.rows[2].cells[1].text = grant.get('number', '')

        table.rows[3].cells[0].text = 'Source'
        table.rows[3].cells[1].text = grant.get('source', '')

        table.rows[4].cells[0].text = 'Acknowledgement'
        table.rows[4].cells[1].text = grant.get('acknowledgement', '')

    def docx_build_person_info(self, person, document, styles):
        table = document.add_table(rows=1, cols=2)
        table.rows[0].cells[0].width = Cm(7.5)
        table.rows[0].cells[0].height = Cm(20)
        table.rows[0].cells[1].width = Cm(11)
        table.rows[0].cells[1].height = Cm(20)

        p = table.rows[0].cells[0].paragraphs[0]
        file_path = self.download_photo_person(person)
        if file_path:
            p.add_run().add_picture(file_path, width=Cm(7))

        p = table.rows[0].cells[1].paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        file_path = self.download_photo_logo()
        if file_path:
            p.add_run().add_picture(file_path, width=Cm(5))

        p = table.rows[0].cells[1].add_paragraph()
        p.style = styles['title_name']
        name_full = person.get('name', '') + ' ' + person.get('surname', '')
        p.add_run('\n' + name_full).bold = True

        table = table.rows[0].cells[1].add_table(rows=7, cols=2)
        table.rows[0].cells[0].width = Cm(3)
        table.rows[0].cells[1].width = Cm(8)

        table.rows[0].cells[0].text = 'Year of birth'
        table.rows[0].cells[1].text = str(person.get('birth', ''))

        table.rows[1].cells[0].text = 'Degree'
        table.rows[1].cells[1].text = person.get('degree', '')

        table.rows[2].cells[0].text = 'Position'
        table.rows[2].cells[1].text = person.get('position', '')

        table.rows[3].cells[0].text = 'Personal page'
        p = table.rows[3].cells[1].paragraphs[0]
        add_hyperlink(p, 'Skoltech profile', person.get('page'))

        table.rows[4].cells[0].text = 'H-index Scholar'
        p = table.rows[4].cells[1].paragraphs[0]
        p.add_run(str(person['h_sch'])).bold = True
        p.add_run('\t[')
        add_hyperlink(p, 'link', person.get('link_sch'))
        p.add_run(']')

        table.rows[5].cells[0].text = 'H-index Scopus'
        p = table.rows[5].cells[1].paragraphs[0]
        p.add_run(str(person['h_sco'])).bold = True
        p.add_run('\t[')
        add_hyperlink(p, 'link', person.get('link_sco'))
        p.add_run(']')

        table.rows[6].cells[0].text = 'H-index WoS'
        p = table.rows[6].cells[1].paragraphs[0]
        p.add_run(str(person['h_wos'])).bold = True
        p.add_run('\t[')
        add_hyperlink(p, 'link', person.get('link_wos'))
        p.add_run(']')

    def docx_build_person_stat(self, stat, document, styles):
        table = document.add_table(rows=5, cols=7)
        table.rows[0].cells[0].width = Cm(5)
        table.style = 'Table Grid'

        for i, year in enumerate(YEARS, 1):
            p = table.rows[0].cells[i].paragraphs[0]
            p.style = styles['table_stat_title']
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(str(year))
        p = table.rows[0].cells[6].paragraphs[0]
        p.style = styles['table_stat_title']
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run('Total').bold = True

        p = table.rows[1].cells[0].paragraphs[0]
        p.style = styles['table_stat_title']
        p.add_run('Q1 journals')
        p = table.rows[2].cells[0].paragraphs[0]
        p.style = styles['table_stat_title']
        p.add_run('Q2 journals')
        p = table.rows[3].cells[0].paragraphs[0]
        p.style = styles['table_stat_title']
        p.add_run('Other journals')
        p = table.rows[4].cells[0].paragraphs[0]
        p.style = styles['table_stat_title']
        p.add_run('Total').bold = True

        for i, year in enumerate(YEARS, 1):
            p = table.rows[1].cells[i].paragraphs[0]
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(str(stat[year]['q1']))
            p = table.rows[2].cells[i].paragraphs[0]
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(str(stat[year]['q2']))
            p = table.rows[3].cells[i].paragraphs[0]
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(str(stat[year]['q0']))
            p = table.rows[4].cells[i].paragraphs[0]
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(str(stat[year]['total']))

        p = table.rows[1].cells[6].paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(str(stat['total']['q1']))
        p = table.rows[2].cells[6].paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(str(stat['total']['q2']))
        p = table.rows[3].cells[6].paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(str(stat['total']['q0']))
        p = table.rows[4].cells[6].paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(str(stat['total']['total']))

    def docx_save(self, document, fpath):
        for section in document.sections:
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(2)
            section.right_margin = Cm(1)

        document.save(fpath)

    def download_excel(self):
        if DEBUG:
            return './_data/cait.xlsx'

        file_path = self.folder + '/' + 'cait.xlsx'

        url = 'https://docs.google.com/spreadsheets/d/1jz76t2bRMzlNqL315SUf1WKr45lWu-c_/edit?usp=sharing&ouid=102021586566196668105&rtpof=true&sd=true'

        uid = url.split('/')[-2]
        download_file_from_google_drive(uid, file_path)

        self.log('Excel file is downloaded', 'res')
        return file_path

    def download_photo_logo(self):
        if DEBUG:
            return f'./_data/cait.jpg'

        file_path = self.folder + '/' + 'cait.jpg'

        url = 'https://drive.google.com/file/d/1hApCr3FnpZedkaJnQkRA4GRoeKY-HZce/view?usp=sharing'

        uid = url.split('/')[-2]
        download_file_from_google_drive(uid, file_path)

        self.log('Logo photo is downloaded', 'res')
        return file_path

    def download_photo_person(self, person):
        if DEBUG:
            return f'./_data/{person["id"][1:]}.jpg'

        file_name = person['id'][1:] + '.jpg'
        file_path = self.folder + '/' + file_name

        url = person['photo']
        if not url:
            return

        uid = url.split('/')[-2]
        download_file_from_google_drive(uid, file_path)

        self.log('Photo of the person is downloaded', 'res')
        return file_path

    def export_word_cv(self):
        if len(self.task.get('authors', [])) != 1:
            text = 'export_word_cv (task should contain only one author)'
            self.log(text, 'err')
            return

        person = self.team.get(self.task['authors'][0])
        if person is None:
            text = 'export_word_cv (invalid team member uid in task)'
            self.log(text, 'err')
            return

        stat = self.get_papers_stat(person['id'], YEARS)

        document, styles = self.docx_create()

        self.docx_build_person_info(person, document, styles)
        document.add_paragraph('\n\n')
        self.docx_build_person_stat(stat, document, styles)
        document.add_paragraph('\n\n')
        text = 'This document is generated automatically. '
        text += 'Journal ratings were determined based on the '
        text += 'Scimago Journal & Country Rank. All publications '
        text += 'in the collected database will be manually checked later.'
        p = document.add_paragraph(text, style='Intense Quote')
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        document.add_page_break()

        document.add_heading('Publications', level=1)
        self.docx_build_paper_list(document, stat, person['id'])

        fname = 'CAIT_' + person['surname'] + '_' + person['name'] + '.docx'
        fpath = self.get_fpath(fname)
        self.docx_save(document, fpath)
        self.log(f'Document "{fpath}" is saved', 'res')

    def export_grant_papers(self):
        if len(self.task.get('grants', [])) != 1:
            text = 'export_grant_papers (task should contain only one grant)'
            self.log(text, 'err')
            return

        grant = self.grants.get(self.task['grants'][0])
        if grant is None:
            text = 'export_grant_papers (invalid grant uid in task)'
            self.log(text, 'err')
            return

        uid = grant['id']

        stat = self.get_papers_stat(years=YEARS, grant=uid)

        document, styles = self.docx_create()

        self.docx_build_grant_info(grant, document, styles)

        document.add_paragraph('\n\n')
        text = 'This document is generated automatically. '
        text += 'Journal ratings were determined based on the '
        text += 'Scimago Journal & Country Rank. All publications '
        text += 'in the collected database will be manually checked later. '
        text += '\n\nNOTE! We use bold type for all authors associated with our center. This will be clarified later (only authors participating in the corresponding grant will be displayed in bold).'
        p = document.add_paragraph(text, style='Intense Quote')
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        document.add_page_break()

        document.add_heading('Publications', level=1)
        self.docx_build_paper_list(document, stat, grant=uid, with_links=True)

        fname = 'CAIT_' + uid[1:] + '.docx'
        fpath = self.get_fpath(fname)
        self.docx_save(document, fpath)
        self.log(f'Document "{fpath}" is saved', 'res')

    def get_journal(self, title=None, issn=None, dist_max=0, dist_max_wrn=1):
        if not issn and (not title or len(title) < 2):
            return

        for title_real, item in self.journals_ref.items():
            t = title_real.lower()
            if issn and issn == item['issn']:
                if title and distance(title.lower(), t) > dist_max:
                    text = 'Journal found by ISSN but titles are different: '
                    text += f'"{title}" is replaced by "{title_real}"'
                    self.log(text, 'wrn')
                return item
            dist = distance(title.lower(), t)
            if dist > dist_max:
                continue
            if dist >= dist_max_wrn:
                text = f'Journal "{title}" is replaced by "{title_real}"'
                self.log(text, 'wrn')
            return item

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
                journal = self.journals[paper['journal']]
                q1 = journal.get('sjr_q1', '')
                q2 = journal.get('sjr_q2', '')
                if q == 1 and len(q1) < 2:
                    continue
                if q == 2 and (len(q1) >= 2 or len(q2) < 2):
                    continue
                if q == 0 and (len(q1) >= 2 or len(q2) >= 2):
                    continue

            res[title] = paper

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

    def get_fpath(self, fname):
        return self.folder + '/' + fname

    def init(self, use_last=False):
        if use_last:
            files = [f for f in os.listdir('./') if f.startswith('export_')]
            files.sort()
            if len(files) == 0:
                self.log('Can not find the last "export_*" folder', 'err')
            self.folder = './' + files[-1]
            self.log(f'The existing folder "{self.folder}" will be used', 'res')

            for f in os.listdir(self.folder):
                if f == 'cait.xlsx':
                    continue
                fpath = self.folder + '/' + f
                if os.path.isfile(fpath):
                    os.remove(fpath)
                else:
                    shutil.rmtree(fpath)
            self.log(f'All files, except DB, are removed from folder', 'res')

            file_path = self.folder + '/' + 'cait.xlsx'
        else:
            self.folder = './export_'
            self.folder += datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            os.mkdir(self.folder)
            self.log(f'The folder "{self.folder}" is created', 'res')

            file_path = self.download_excel()

        self.wb = openpyxl.load_workbook(file_path)
        self.log('Excel file is opened', 'res')

    def load_journals_ref(self):
        res = {}

        def parse_issn(issn):
            if not issn or len(issn) < 8:
                return ''
            return issn[-8:-4] + '-' + issn[-4:]

        def parse_quartiles(data):
            res = {}
            for item in data.split('; '):
                name = item
                quartile = ''
                if item.endswith('(Q1)') or item.endswith('(Q2)') or item.endswith('(Q3)') or item.endswith('(Q4)'):
                    name = item[:-5]
                    quartile = item[-3:-1]
                res[name] = quartile
            return res

        def parse_quartiles_q(quartiles, kind='Q1'):
            res = []
            for name, item in quartiles.items():
                if kind == 'Q0' and not item or kind == item:
                    res.append(name)
            return ', '.join(res)

        with open(JOURNALS_REF_SJR_FILE_PATH, newline='') as f:
            for i, row in enumerate(csv.reader(f, delimiter=';')):
                if i==0: continue
                title = row[2]
                quartiles = parse_quartiles(row[19])
                res[title] = {
                    'title': title,
                    'type': row[3],
                    'issn': parse_issn(row[4]),
                    'country': row[15],
                    'publisher': row[17],
                    'sjr_rank': row[5],
                    'sjr_best_quartile': row[6],
                    'sjr_impact': row[7],
                    'sjr_quartiles': quartiles,
                    'sjr_q1': parse_quartiles_q(quartiles, 'Q1'),
                    'sjr_q2': parse_quartiles_q(quartiles, 'Q2'),
                    'sjr_q3': parse_quartiles_q(quartiles, 'Q3'),
                    'sjr_q4': parse_quartiles_q(quartiles, 'Q4'),
                    'sjr_q0': parse_quartiles_q(quartiles, 'Q0'),
                }

        return res

    def load_table(self, name, max_count=10000):
        sh = self.wb[name]

        i = 1
        fields = []
        for j in range(1, max_count):
            field = sh.cell(i, j).value
            if field is None or not field or field == ' ':
                break

            field = field.lower()
            field = field.replace(' ', '_')
            fields.append(field)

        table = {}
        for i in range(2, max_count):
            uid = sh.cell(i, 1).value
            if uid is None or not uid or uid == ' ':
                break

            row = {}
            for j, field in enumerate(fields, 1):
                value = sh.cell(i, j).value
                if value is not None and value != ' ':
                    row[field] = value

            table[uid] = row

        return table

    def log(self, text, kind='err'):
        res = 'IBM '
        res += f'[{kind.upper()}] '
        res += '>>> '
        res += text
        print(res)
        if kind == 'err':
            self.log('The system will shut down due to an error', 'wrn')
            sys.exit(0)


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a w:r element and a new w:rPr element
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Create a new Run object and add the hyperlink into it
    r = paragraph.add_run ()
    r._r.append (hyperlink)

    r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True

    return hyperlink


def download_file_from_google_drive(id, destination):
    URL = "https://docs.google.com/uc?export=download"

    def get_confirm_token(response):
        for key, value in response.cookies.items():
            if key.startswith('download_warning'):
                return value

        return None

    def save_response_content(response, destination):
        CHUNK_SIZE = 32768

        with open(destination, "wb") as f:
            for chunk in response.iter_content(CHUNK_SIZE):
                if chunk: # filter out keep-alive new chunks
                    f.write(chunk)

    session = requests.Session()

    response = session.get(URL, params = { 'id' : id }, stream = True)
    token = get_confirm_token(response)

    if token:
        params = { 'id' : id, 'confirm' : token }
        response = session.get(URL, params = params, stream = True)

    save_response_content(response, destination)


if __name__ == '__main__':
    use_last = len(sys.argv) > 1 and sys.argv[1] in ['-l', '--last']
    ibm = InBiMa(use_last)

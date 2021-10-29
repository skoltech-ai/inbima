import os
from pybtex.database import BibliographyDataError
from pybtex.database import parse_string
import sys
import xlsxwriter


from inbima import InBiMa


YEAR_MIN = 2017
BIB_FOLDER_PATH = './bib_wos/'


def parse_author(data, team={}):
    authors = ''
    authors_parsed = ''

    for author_object in data:
        author = str(author_object)
        author = author.replace(',', '')
        author = author.split(' ')[:2]
        surname, name = author if len(author) == 2 else [author, '']
        name = name.replace('.', '')
        name_char = name[0] if name else None

        if name_char:
            author = surname + ' ' + name_char + '.' + ', '
        else:
            author = surname + ', '

        uid_found = None
        for uid, item in team.items():
            if surname.lower() == item['surname'].lower():
                if len(name) >= 2 and name.lower() != item['name'].lower():
                    print(f'> !!! Surnames "{surname}" are equal but names ("{name}" vs "{item["name"]}") not. Skip')
                    continue
                uid_found = uid
                break

        authors += author
        authors_parsed += uid_found + ', ' if uid_found else author

    if not '#' in authors_parsed:
        print(f'> !!! Not found team members for "{authors}"')

    return authors[:-2], authors_parsed[:-2]


def parse_grant(grant_str, grants={}):
    if not grant_str or len(grant_str) < 5:
        return

    grant = []
    for uid, item in grants.items():
        if item['number'].lower() in grant_str.lower():
            grant.append(uid)
    return ', '.join(grant)


def load_bibs(uids):
    bibs = None

    for uid in uids:
        bib_file = BIB_FOLDER_PATH + uid + '.bib'
        print(f'> ... Parse bib file : ', bib_file)

        with open(bib_file, 'r') as f:
            data = f.read()
        data = data.replace('Early Access Date', 'EarlyAccessDate')

        if bibs is None:
            bibs = parse_string(data, bib_format='bibtex')
        else:
            bibs_ = parse_string(data, bib_format='bibtex')

            for key, entry in bibs_.entries.items():
                try:
                    bibs.add_entry(key, entry)
                except BibliographyDataError:
                    pass # Duplicated entry

    return bibs


def run(uids):
    ibm = InBiMa()

    bibs = load_bibs(uids)

    wb = xlsxwriter.Workbook(f'./parser_wos_result.xlsx')
    ws1 = wb.add_worksheet('papers_parsed')
    ws2 = wb.add_worksheet('papers_conf_parsed')
    ws3 = wb.add_worksheet('journals_parsed')
    ws4 = wb.add_worksheet('fields_parsed')

    for ws in [ws1, ws2]:
        ws.write(0, 0, 'Title')
        ws.write(0, 1, 'Year')
        ws.write(0, 2, 'Authors')
        ws.write(0, 3, 'Journal')
        ws.write(0, 4, 'Volume')
        ws.write(0, 5, 'Number')
        ws.write(0, 6, 'Pages')
        ws.write(0, 7, 'Site')
        ws.write(0, 8, 'PDF')
        ws.write(0, 9, 'Screen')
        ws.write(0, 10, 'DOI')
        ws.write(0, 11, 'Grant')
        ws.write(0, 12, 'Grant_str')
        ws.write(0, 13, 'Authors Parsed')
        ws.write(0, 14, 'Note')

    ind_paper = 0
    ind_conf = 0
    journals = {}
    for tag, bib in (bibs.entries.items() if bibs else []):
        is_paper_conf = False

        title = (bib.fields.get('title') or '{}')[1:-1]
        if not title:
            print(f'> !!! No title for paper "{tag}"')
            continue
        title = title.replace('\n', ' ')

        year = (bib.fields.get('year') or '{}')[1:-1]
        if not year:
            print(f'> !!! No year for paper "{title}"')
            continue
        if int(year) < YEAR_MIN:
            continue

        authors, authors_parsed = parse_author(bib.persons['author'], ibm.team)
        if not authors:
            print(f'> !!! No authors for paper "{title}"')
            continue

        issn = (bib.fields.get('issn') or '{}')[1:-1]
        journal = (bib.fields.get('journal') or '{}')[1:-1]
        if not journal:
            journal = (bib.fields.get('booktitle') or '{}')[1:-1]
            if not journal:
                print(f'> !!! No journal/booktitle for paper "{title}"')
                continue
            is_paper_conf = True
            journal = journal.replace('\n', ' ')
        else:
            journal_object = ibm.get_journal(journal, issn)
            if journal_object:
                journal = journal_object['title']
                journals[journal] = journal_object
            else:
                print(f'> ! Journal "{journal.lower()}" is not recognized for paper "{title}"')
                journals[journal] = {'title': journal, 'issn': issn}

        volume = (bib.fields.get('volume') or '{ }')[1:-1]
        number = (bib.fields.get('number') or '{ }')[1:-1]
        pages = (bib.fields.get('pages') or '{ }')[1:-1]
        grant_str = (bib.fields.get('Funding-Text') or '{ }')[1:-1]
        grant = parse_grant(grant_str, ibm.grants) or ' '

        note = ' '
        if is_paper_conf:
            note += (bib.fields.get('series') or '{ }')[1:-1] + ' '
            note += (bib.fields.get('note') or '{ }')[1:-1] + ' '
            note += (bib.fields.get('organization') or '{ }')[1:-1] + ' '

        if is_paper_conf:
            ind_conf += 1
            ind = ind_conf
            ws = ws2
        else:
            ind_paper += 1
            ind = ind_paper
            ws = ws1

        ws.write(ind, 0, title)
        ws.write(ind, 1, year)
        ws.write(ind, 2, authors)
        ws.write(ind, 3, journal)
        ws.write(ind, 4, volume)
        ws.write(ind, 5, number)
        ws.write(ind, 6, pages)
        ws.write(ind, 7, ' ')
        ws.write(ind, 8, ' ')
        ws.write(ind, 9, ' ')
        ws.write(ind, 10, ' ')
        ws.write(ind, 11, grant)
        ws.write(ind, 12, grant_str)
        ws.write(ind, 13, authors_parsed)
        ws.write(ind, 14, note)

    ws = ws3
    ws.write(0, 0, 'Title')
    ws.write(0, 1, 'ISSN')
    ws.write(0, 2, 'Country')
    ws.write(0, 3, 'Publisher')
    ws.write(0, 4, 'SJR Rank')
    ws.write(0, 5, 'SJR Impact')
    ws.write(0, 6, 'SJR Q1')
    ws.write(0, 7, 'SJR Q2')
    ws.write(0, 8, 'SJR Q3')
    ws.write(0, 9, 'SJR Q4')
    ws.write(0, 10, 'Note')

    ind = 0
    for item in journals.values():
        ind += 1
        ws.write(ind, 0, item.get('title', ' '))
        ws.write(ind, 1, item.get('issn', ' '))
        ws.write(ind, 2, item.get('country', ' '))
        ws.write(ind, 3, item.get('publisher', ' '))
        ws.write(ind, 4, item.get('sjr_rank', ' '))
        ws.write(ind, 5, item.get('sjr_impact', ' '))
        ws.write(ind, 6, item.get('sjr_q1', ' '))
        ws.write(ind, 7, item.get('sjr_q2', ' '))
        ws.write(ind, 8, item.get('sjr_q3', ' '))
        ws.write(ind, 9, item.get('sjr_q4', ' '))
        ws.write(ind, 10, item.get('note', ' '))

    quartile_names = []
    for item in journals.values():
        for quartile_name in item.get('sjr_quartiles', {}).keys():
            if not quartile_name in quartile_names:
                quartile_names.append(quartile_name)
    quartile_names.sort()

    ws = ws4
    ws.write(0, 0, 'SJR quartile fields')

    ind = 0
    for quartile_name in quartile_names:
        ind += 1
        ws.write(ind, 0, quartile_name)

    wb.close()


if __name__ == '__main__':
    if len(sys.argv) > 1:
        uids = [sys.argv[1]]
    else:
        files = os.listdir(BIB_FOLDER_PATH)
        uids = [f.split('.')[0] for f in files if f.endswith('.bib')]
        uids.sort()

    run(uids)

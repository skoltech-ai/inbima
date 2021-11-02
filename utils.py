import sys


def load_table(sh, max_count=10000):
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


def log(text, kind='err'):
    res = 'IBM '
    res += f'[{kind.upper()}] '
    res += '>>> '
    res += text
    print(res)
    if kind == 'err':
        log('The system will shut down due to an error', 'wrn')
        sys.exit(0)

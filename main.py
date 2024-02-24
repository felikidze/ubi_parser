import openpyxl


book = openpyxl.load_workbook("thrlist.xlsx")
sheet = book.active

potential_map = dict.fromkeys([
    'Внешний нарушитель с низким потенциалом',
    'Внешний нарушитель со средним потенциалом',
    'Внешний нарушитель с высоким потенциалом',
    'Внутренний нарушитель с низким потенциалом',
    'Внутренний нарушитель со средним потенциалом',
    'Внутренний нарушитель с высоким потенциалом'
], None)


def proxy_set_dict_value(key, field, value):
    if not potential_map[key].get(field):
        potential_map[key][field] = ""

    if isinstance(value, list):
        for item in value.split(','):
            if item.lower() not in potential_map[key].get(field).lower():
                potential_map[source_trim][field] += f"{item}; "
    else:
        if value.lower() not in potential_map[key].get(field).lower():
            potential_map[source_trim][field] += f"{value}; "


for row in range(3, sheet.max_row):
    name = sheet[row][1].value
    sources = sheet[row][3].value
    targets = sheet[row][4].value

    for source in sources.split(';'):
        source_trim = source.strip()
        if source_trim == '':
            continue

        if not potential_map.get(source_trim):
            potential_map[source_trim] = {}

        for field, value in {'target': targets, 'name': name}.items():
            proxy_set_dict_value(source_trim, field, value)

with open("data.txt", "w") as file:
    for item in potential_map.items():
        file.write(f"{item[0]}\n")
        file.write(f"Объекты воздействия: {item[1]['target']}\n")
        file.write(f"Наименования УБИ: {item[1]['name']}\n\n")
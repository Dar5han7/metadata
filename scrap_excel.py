import csv

input_csv_file_name = "NGMA_DEL_10-09-2020_ANAND.csv"
output_csv_file_name = "NGMA_DEL_10-09-2020_ANAND_formatted.csv"
title_word = "Title"
tables = []

with open(input_csv_file_name, 'r' ) as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    table = []
    for row in csv_reader:
        if row[0] == title_word:
            tables.append(table)
            table = []
        table.append(row)
    tables.append(table)

tables = tables[1:]
for table in tables:
    for row in table:
        print(row)
    print()

number_of_tables = len(tables)
parsed = {}

for table_number, table in enumerate(tables):
    for row in table:
        key = row[0]
        value = row[1]
        try:
            parsed[key][table_number] = value
        except:
            parsed[key] = ["" for x in range(number_of_tables)]
            parsed[key][table_number] = value

with open(output_csv_file_name, 'w' ) as f:
    csv_writer = csv.writer(f)

    for key, value in parsed.items():
        row = [key] + value
        print(row)
        csv_writer.writerow(row)
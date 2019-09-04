import xlrd
from datetime import datetime

records = []

spreadsheet = xlrd.open_workbook(
    r'C:\\Users\\MDL\\SedgwickTXT.xlsx'
).sheet_by_index(0)


headers = []
for col in range(spreadsheet.ncols):
    headers.append(spreadsheet.cell_value(0, col))
for row in range(1, spreadsheet.nrows):
    record = {}

    for col in range(spreadsheet.ncols):
        record[headers[col]] = spreadsheet.cell_value(
            row, col
        )
    records.append(record)


date_string_format = "%Y-%m-%dT%H:%M:%S"

ascending_sorted_records = sorted(
    records,
    key=lambda x: datetime.strptime(
        x["dateopened"], date_string_format
    )
)

print(ascending_sorted_records)[0]

counts = {}


def add_to_count(key):
    if key in counts:
            counts[key] += 1
    else:
        counts[key] = 1
    return


def count_open_claims(occupations=[]):
    if len(occupations) == 0:
        claims = ascending_sorted_records
    else:
        claims = list(
            filter(
                lambda record: record["occuption"]
                in occupations,
                ascending_sorted_records,
            )
        )
    for claim in claims:
        date_opened = datetime.strptime(
            claim["dateopened"], date_string_format
        )

        year_opened = date_opened.strftime("%Y")
        month_opened = date_opened.strftime("%m")
        key = f"{month_opened}_{year_opened}"

        # Case 1
        has_no_close_date = claim["dateclosed"] == "null"

        if has_no_close_date:
            add_to_count(key)
            continue

        # Case 2
        date_closed = datetime.strptime(
            claim["dateclosed"], date_string_format
        )

        if claim["datereopened"] != "null":
            date_reopened = datetime.strptime(
                claim["datereopened"], date_string_format
            )
            has_reopened_this_month = (
                date_reopened >= date_closed
            )

            if has_reopened_this_month:
                add_to_count(key)
                continue
        
        # Case 3
        year_closed = date_closed.strftime("%Y")
        month_closed = date_closed.strftime("%m")

        # If there's no reopen date, but a close date exists, we just need to make sure closing month/year is not the same as opening month/year.
        did_close_in_opening_month = (
            key == f"{month_closed}_{year_closed}"
        )

        if did_close_in_opening_month:
            add_to_count(key)
            continue
    return counts


result = count_open_claims()
# print(result['01_2018'])
print(result)
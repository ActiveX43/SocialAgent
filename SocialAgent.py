import gspread
import datetime
import holidays
import calendar
import math
import time
import json

class DailySheet:
    row_num = 7
    index_num = 3
    index = ['본관', '민원실', '정문']
    #                  0    1    2    3     4     5
    morning_worker = ['1', '3', '8', '8', '13', '주']
    #              0    1    2    3    4    5    6    7     8     9    10    11    12    13    14    15 
    all_worker = ['1', '2', '3', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17']

    def __init__(self, day):
        self.date = day
        "-----Modify Time Table Structure-----"
        self.tt_struct = [
                {'start': 1, "num": 4},
                {'start': 0, "num": 7},
                {'start': 1, "num": 6}
            ] 
        "-------------------------------------"
        self.worker_idx = -1
        self.morning_idx = -1
        self.rotation = []


class MonthlySheet(DailySheet):
    def __init__(self, day, exception):
        super().__init__(day)
        self.office_day = []
        holiday = [x[0].strftime("%Y-%m-%d") for x in holidays.KR(years=day.year).items() \
                if x[0].month == day.month]
        _, day_count = calendar.monthrange(day.year, day.month)
        for day_it in range(1, day_count+1):
            d = datetime.date(day.year, day.month, day_it)
            if d.weekday() < 5 \
            and d.strftime("%Y-%m-%d") not in holiday + exception:
                today_sheet = DailySheet(d)
                self.office_day.append(today_sheet)

    def set_monthly_time(self, worker_idx, morning_idx):
        cur_worker_idx = worker_idx
        cur_morning_idx = morning_idx
        """------Modify Work Rotation------"""
        for day in self.office_day:
            day.worker_idx = cur_worker_idx
            day.morning_idx = cur_morning_idx

            rot = day.all_worker[cur_worker_idx:] + day.all_worker[:cur_worker_idx]
            day.rotation = rot[:day.tt_struct[0]['num']] + [day.morning_worker[cur_morning_idx]] + rot[day.tt_struct[0]['num']:]

            cur_worker_idx = (cur_worker_idx - 1) % len(self.all_worker)
            cur_morning_idx = (cur_morning_idx + 1) % len(self.morning_worker)
        """--------------------------------"""


def num2cell(cell1, cell2=None):
    if cell2 is None:
        cell_str = f"{chr(cell1[1] + ord('A') - 1)}{cell1[0]}"
        return cell_str
    else:
        cell_str = f"{chr(cell1[1]+ ord('A') - 1)}{cell1[0]}:{chr(cell2[1]+ ord('A') - 1)}{cell2[0]}"
        return cell_str


def keep_try(original_function):
    def wrapper(*args, **kwargs):
        finished = False
        while not finished:
            try:
                res = original_function(*args, **kwargs)
            except gspread.exceptions.APIError as e:
                code = str(e)[9:12]
                if code == "429":
                    print(f"{code}:Too many request")
                    time.sleep(10)
                else:
                    print(e)
                    exit()
            else:
                finished = True
        return res

    return wrapper


@keep_try    
def init_sheets(auth, filename):
    gc = gspread.service_account(filename=auth)
    return gc.open(filename)

@keep_try
def cell_sizing(ss, sht, loc, rg, pixel):
    sheet_id = sht._properties['sheetId']
    if loc != "COLUMNS" and loc != "ROWS":
        return
    body = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": loc,
                            "startIndex": rg[0]-1,
                            "endIndex": rg[1]
                            },
                        "properties": {
                            "pixelSize": pixel
                            },
                        "fields": "pixelSize"
                        }
                    }
                ]
            }

    return ss.batch_update(body)

def cell_styling(ss, sht, i, today):
    sheet_id = sht._properties['sheetId']
    """
                      cur_col                    next_col 
               ______ ______ ______ ______ ______ ______
              |      |______|______|______|      |      |
     cur_row->|      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |______|______|______|______|
              |
    next_row->|
    """
    cur_row = i//3*(today.row_num + 1) + 3
    next_row = (i//3+1)*(today.row_num + 1) + 3
    cur_col = i%3*(today.index_num + 1) + 3
    next_col = (i%3+1)*(today.index_num + 1) + 3

    body1 = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": cur_col-2,
                            "endIndex": next_col-2
                            },
                        "properties": {
                            "pixelSize": 80
                            },
                        "fields": "pixelSize"
                        }
                    
                    }
                ]
            }

    body2 = {
            "requests": [
                {
                    "updateCells": {
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredFormat": {
                                            "horizontalAlignment": "CENTER",
                                            "verticalAlignment": "MIDDLE"
                                            }
                                        }
                                    ]
                                }
                            ],
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": cur_row-2,
                            "endRowIndex": next_row-2,
                            "startColumnIndex": cur_col-2,
                            "endColumnIndex": next_col-2
                            },
                        "fields": "userEnteredFormat"
                        }
                    }
                ]
            }

    body3 = {
            "requests": [
                {
                    "updateBorders": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": cur_row-2,
                            "endRowIndex": next_row-2,
                            "startColumnIndex": cur_col-2,
                            "endColumnIndex": next_col-2
                            },
                        "top": {
                            "style": "SOLID_MEDIUM"
                            },
                        "bottom": {
                            "style": "SOLID_MEDIUM"
                            },
                        "left": {
                            "style": "SOLID_MEDIUM"
                            },
                        "right": {
                            "style": "SOLID_MEDIUM"
                            },
                        "innerHorizontal": {
                            "style": "SOLID"
                            },
                        "innerVertical": {
                            "style": "SOLID"
                            }
                        }
                    }
                ]
            }


    try_batch_update = keep_try(ss.batch_update)
    try_batch_update(body1)
    try_batch_update(body2)
    try_batch_update(body3)


@keep_try
def cell_data(sht, i, today):
    """
                      cur_col                    next_col 
               ______ ______ ______ ______ ______ ______
              |      |______|______|______|      |      |
     cur_row->|      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |______|______|______|______|
              |
    next_row->|
    """
    cur_row = i//3*(today.row_num + 1) + 3
    next_row = (i//3+1)*(today.row_num + 1) + 3
    cur_col = i%3*(today.index_num + 1) + 3
    next_col = (i%3+1)*(today.index_num + 1) + 3


    #day_cell = num2cell((cur_row-1, cur_col-1))
    #sht.update(day_cell, today.date.strftime("%m/%d"))

    #index_cells = num2cell((cur_row-1, cur_col), (cur_row-1, cur_col+today.index_num-1))
    #sht.update(index_cells, [today.index])

    rotation_cells = num2cell((cur_row-1, cur_col-1), (next_row-2, next_col-2))
    time_table = []
    pg = 0
    for col in range(today.index_num):
        time_table_col = []
        time_table_col += [""]*(today.tt_struct[col]['start'])
        time_table_col += today.rotation[pg:(pg + today.tt_struct[col]['num'])]
        time_table_col += [""]*(today.row_num - today.tt_struct[col]['start'] - today.tt_struct[col]['num'])
        time_table.append(time_table_col)
        pg += today.tt_struct[col]['num']
    time_table = [[""]*(today.row_num)] + time_table
    time_table = list(map(list,zip(*time_table)))
    time_table = [[today.date.strftime("%m/%d")] + today.index] + time_table
    sht.update(rotation_cells, time_table)

@keep_try
def cell_merge(sht, i, today):
    """
                      cur_col                    next_col 
               ______ ______ ______ ______ ______ ______
              |      |______|______|______|      |      |
     cur_row->|      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |      |______|______|______|
              |______|______|______|______|
              |
    next_row->|
    """
    cur_row = i//3*(today.row_num + 1) + 3
    next_row = (i//3+1)*(today.row_num + 1) + 3
    cur_col = i%3*(today.index_num + 1) + 3

    merged_cells = num2cell((cur_row-1, cur_col-1), (next_row-2, cur_col-1))
    sht.merge_cells(merged_cells)


def create_sheet(ss, now_day, first_idx, exception):
    now = MonthlySheet(now_day, exception)
    now.set_monthly_time(first_idx[0], first_idx[1])

    row_num = max(70, math.ceil(len(now.office_day)/3)*(now.row_num + 1) + 6)
    col_num = 3*(now.index_num + 1) + 4

    sht = ss.add_worksheet(title=now_day.strftime("%m월"), rows=str(row_num), cols=str(col_num))

    cell_sizing(ss, sht, "COLUMNS", (1, 1), 20)

    for i, today in enumerate(now.office_day):
        cell_merge(sht, i, today)
        cell_data(sht, i, today)
        cell_styling(ss, sht, i, today)


if __name__ == "__main__":
    filename = "청원경찰서 공익 근무표"
    ss = init_sheets("test.json", filename) 

    day = datetime.date(2022, 6, 1)

    #  0   1   2   3   4   5   6   7    8    9   10   11   12   13   14   15
    # '1','2','3','5','6','7','8','9','10','11','12','13','14','15','16','17'
    #  0   1   2   3    4   5
    # '1','3','8','8','13','주'
    #           ord, morn
    first_idx = (10, 3) 
    #          "2022-06-01"
    exception = [
            "2022-06-01"
            ]
    create_sheet(ss, day, first_idx, exception)

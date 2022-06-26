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
    index = ['민원실', '정문1', "정문2"]
    #                  0    1    2    3     4     5
    morning_worker = ['1', '3', '8', '8', '13', '주']
    #              0    1    2    3    4    5    6    7     8     9    10    11    12    13    14    15 
    all_worker = ['1', '2', '3', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18']

    def __init__(self, start_day, end_day=None):
        self.date = start_day
        self.end_day = end_day
        "-----Modify Time Table Structure-----"
        self.tt_struct = [
                {'start': 0, "num": 7},
                {'start': 1, "num": 5},
                {'start': 1, "num": 5}
            ] 
        "-------------------------------------"
        self.worker_idx = -1
        self.morning_idx = -1
        self.rotation = []


class MonthlySheet(DailySheet):
    def __init__(self, exception, start_day, end_day=None):
        super().__init__(start_day, end_day)
        self.office_day = []
        holiday = [x[0].strftime("%Y-%m-%d") for x in holidays.KR(years=self.date.year).items() \
                if x[0].month == self.date.month]
        if end_day is None:
            _, end_day_temp = calendar.monthrange(self.date.year, self.date.month)
        else:
            end_day_temp = end_day.day

        for day_it in range(start_day.day, end_day_temp+1):
            d = datetime.date(self.date.year, self.date.month, day_it)
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
            del rot[day.row_num-1]
            day.rotation = [day.morning_worker[cur_morning_idx]] + rot
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


def access_sheet(ss, start_day):
    row_num = 70
    col_num = 17
    try:
        sht = ss.worksheet(start_day.strftime("%m월"))
        return sht
    except gspread.exceptions.WorksheetNotFound:
        sht = ss.add_worksheet(title=start_day.strftime("%m월"), rows=str(row_num), cols=str(col_num))
        return sht


def fill_sheet(ss, sht, first_idx, exception, start_day, end_day=None, start_point=0):
    now = MonthlySheet(exception, start_day, end_day)
    now.set_monthly_time(first_idx[0], first_idx[1])

    cell_sizing(ss, sht, "COLUMNS", (1, 1), 20)

    for i, today in enumerate(now.office_day):
        cell_merge(sht, i+start_point, today)
        cell_data(sht, i+start_point, today)
        cell_styling(ss, sht, i+start_point, today)


if __name__ == "__main__":
    filename = "청원경찰서 공익 근무표"
    ss = init_sheets("test.json", filename) 

    start_day = datetime.date(2022, 7, 1)
    #end_day = datetime.date(2022, 7, 10)
    start_point = 0

    #  0   1   2   3   4   5   6   7    8    9   10   11   12   13   14   15
    # '1','2','3','5','6','7','8','9','10','11','12','13','14','15','16','17'
    #  0   1   2   3    4   5
    # '1','3','8','8','13','주'
    #           ord, morn
    first_idx = (10, 5) 

    #          "2022-06-01"
    exception = [
            ]

    sht = access_sheet(ss, start_day)
    fill_sheet(ss, sht, first_idx, exception, start_day, start_point=start_point)

import pandas as pd
import calendar
import holidays
import string
from collections import defaultdict
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

def run_dutyple(input_path, output_path,
                nurse_count, year, month,
                weekday_D, weekday_E, weekday_N,
                holiday_D, holiday_E, holiday_N,
                N_count_nurse):

    weekday_W = weekday_D + weekday_E
    holiday_W = holiday_D + holiday_E

    if nurse_count > 26:
        raise ValueError("간호사는 최대 26명까지 지원됩니다 (A~Z)")

    nurse = list(string.ascii_uppercase[:nurse_count])
    duty = ["N", "W", "X"]
    weight = {"N": 0, "W": 1, "X": 2}

    _, num_day = calendar.monthrange(year, month)

    def get(year, month):
        month_holiday = [day.day for day in holidays.CountryHoliday('KR', years=year) if day.month == month]
        month_end = {day for day in range(1, num_day + 1) if calendar.weekday(year, month, day) in [5, 6]}
        valid_holend = month_end - set(month_holiday)
        return month_holiday, month_end, valid_holend

    month_holiday, month_end, valid_holend = get(year, month)
    daily_all = [-2, -1, 0] + list(range(1, num_day + 1))
    daily = list(range(1, num_day + 1))
    daily_end = sorted(month_end | set(month_holiday))
    daily_week = [day for day in daily if day not in daily_end]

    daily_wallet = {}
    for day in daily:
        if day in daily_end:
            daily_wallet[day] = {'N': holiday_N, 'W': holiday_W, 'X': holiday_W}
        else:
            daily_wallet[day] = {'N': weekday_N, 'W': weekday_W, 'X': weekday_W}

    nurse_wallet = {}
    for worker in nurse:
        X_count_nurse = len(daily_end) + 1
        W_count_nurse = len(daily) - N_count_nurse - X_count_nurse + 4
        nurse_wallet[f"{worker}_wallet"] = {"N": N_count_nurse, "X": X_count_nurse, "W": W_count_nurse}

    df_origin = pd.read_excel(input_path, index_col=0)
    origin_cols = list(df_origin.columns)
    extra_cols = [-2, -1, 0]
    final_cols = extra_cols + [c for c in origin_cols if c not in extra_cols]
    df = pd.DataFrame(index=df_origin.index, columns=final_cols)

    def prefer(nurse, day, duty):
        if f"{nurse}_wallet" in nurse_wallet:
            if duty == "N" and nurse_wallet[f"{nurse}_wallet"]["N"] > 0:
                nurse_wallet[f"{nurse}_wallet"]["N"] -= 1
            elif duty == "X" and nurse_wallet[f"{nurse}_wallet"]["X"] > 0:
                nurse_wallet[f"{nurse}_wallet"]["X"] -= 1
            elif duty == "W" and nurse_wallet[f"{nurse}_wallet"]["W"] > 0:
                nurse_wallet[f"{nurse}_wallet"]["W"] -= 1
        if day in daily_wallet:
            if daily_wallet[day][duty] > 0:
                daily_wallet[day][duty] -= 1
        df.loc[nurse, day] = duty

    for col in [-2, -1, 0]:
        if col in df_origin.columns:
            for nurse_name in df.index:
                val = df_origin.loc[nurse_name, col]
                if val in ["D", "E"]:
                    df.loc[nurse_name, col] = "W"
                elif val in ["N", "X"]:
                    df.loc[nurse_name, col] = val
                else:
                    df.loc[nurse_name, col] = None

    for nurse_name in df_origin.index:
        for day in df_origin.columns:
            if isinstance(day, int) and day > 0:
                try:
                    raw_val = df_origin.at[nurse_name, day]
                except KeyError:
                    continue  # 열 이름이 잘못되었거나 존재하지 않는 경우 건너뜀
                if pd.notna(raw_val):
                    duty = str(raw_val).strip().upper()
                    if duty in ["D", "E"]:
                        prefer(nurse_name, day, "W")
                    elif duty in ["N", "X"]:
                        prefer(nurse_name, day, duty)


    df["D"] = (df == "D").sum(axis=1)
    df["E"] = (df == "E").sum(axis=1)
    df["N"] = (df == "N").sum(axis=1)
    df["X"] = (df == "X").sum(axis=1)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="dutyple")

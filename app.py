import streamlit as st
import pandas as pd
from calendar import monthrange
from datetime import date
import datetime
import random
import io

# 员工姓名
names = st.text_input("员工姓名, 逗号分隔")

# 选择年份
selected_year = st.selectbox("选择年份", list(range(2020, 2100)), index=3)
# 选择月份
selected_month = st.selectbox("选择月份", list(range(1, 13)))

# 上班时间
morgen_time = st.time_input("早上上班时间", datetime.time(12, 00))
bis_morgen = st.time_input("早上下班时间", datetime.time(15, 00))
# 下午上班时间
mittag_time = st.time_input("下午上班时间", datetime.time(17, 30))
bis_mittag = st.time_input("晚上下班时间", datetime.time(22, 30))


# 每周工作天数， 每天工作小时数
work_days, work_hours = 5, 8

# Define a mapping from English weekdays to German weekdays
weekday_translation = {
    'Monday': 'Montag',
    'Tuesday': 'Dienstag',
    'Wednesday': 'Mittwoch',
    'Thursday': 'Donnerstag',
    'Friday': 'Freitag',
    'Saturday': 'Samstag',
    'Sunday': 'Sonntag'
}

# Translate the weekdays to German


def get_dates_and_weekdays_for_month(year, month):
    # Get the number of days in the selected month
    _, num_days = monthrange(year, month)
    
    # Generate the list of dates and weekdays
    dates = [(date(year, month, day+1)).strftime('%d.%m.%Y') for day in range(num_days)]
    weekdays = [(date(year, month, day+1)).strftime('%A') for day in range(num_days)]
    
    german_weekdays_list = [weekday_translation[day] for day in weekdays]
    
    return dates, german_weekdays_list

dates_list, weekdays_list = get_dates_and_weekdays_for_month(selected_year, selected_month)



def add_work_hours_random_offdays(german_dates, german_weekdays,name):
    von_morgen_list = []
    end_morgen_list = []
    von_nacht_list = []
    end_nacht_list = []
    gesamt_stunden_list = []

    work_days_counter = 0  # To keep track of consecutive work days
    
    # Randomly select the two off days after 5 workdays
    seed_value = sum(ord(char) for char in name)
    random.seed(seed_value)
    off_days = ["Montag"]
    remaining_days = [day for day in weekday_translation.values() if day not in off_days]
    additional_off_day = random.choice(remaining_days)
    off_days.append(additional_off_day)
    
    for weekday in german_weekdays:
        if weekday in off_days:  # If it's an off day
            von_morgen_list.append('')
            end_morgen_list.append('')
            von_nacht_list.append('')
            end_nacht_list.append('')
            gesamt_stunden_list.append('')
        else:  # Work days
            von_morgen_list.append(morgen_time)
            end_morgen_list.append(bis_morgen)
            von_nacht_list.append(mittag_time)
            end_nacht_list.append(bis_mittag)
            gesamt_stunden_list.append(work_hours)

    # for idx, weekday in enumerate(german_weekdays):
    #     if work_days_counter < 5:  # If it's a work day
    #         von_morgen_list.append(morgen_time)
    #         end_morgen_list.append(bis_morgen)
    #         von_nacht_list.append(mittag_time)
    #         end_nacht_list.append(bis_mittag)
    #         gesamt_stunden_list.append(work_hours)
    #         work_days_counter += 1
    #     else:  # Off days
    #         von_morgen_list.append('')
    #         end_morgen_list.append('')
    #         von_nacht_list.append('')
    #         end_nacht_list.append('')
    #         gesamt_stunden_list.append('')
    #         work_days_counter += 1
    #         if work_days_counter == 7:  # Reset counter after 7 days
    #             work_days_counter = 0
        # else:  # Two days off after 5 work days
        #     if idx % 7 in off_days_indices:
        #         von_morgen_list.append('')
        #         end_morgen_list.append('')
        #         von_nacht_list.append('')
        #         end_nacht_list.append('')
        #         gesamt_stunden_list.append('')
        #         work_days_counter += 1
        #         if work_days_counter == 7:  # Reset counter after 7 days
        #             work_days_counter = 0
        #     else:
        #         von_morgen_list.append(morgen_time)
        #         end_morgen_list.append(bis_morgen)
        #         von_nacht_list.append(mittag_time)
        #         end_nacht_list.append(bis_mittag)
        #         gesamt_stunden_list.append(8)
        #         work_days_counter += 1

    return von_morgen_list, end_morgen_list, von_nacht_list, end_nacht_list, gesamt_stunden_list



all_names = [names.strip() for names in names.split(",")]

df_all = []

# construct a dataframe to store the data, save them as excel for every employee
for name in all_names:
    von_morgen_random, end_morgen_random, von_nacht_random, end_nacht_random, gesamt_stunden_random = add_work_hours_random_offdays(dates_list, weekdays_list, name)
    
    df = pd.DataFrame({
        "Datum": dates_list,
        "Wochentag": weekdays_list,
        "Von.": von_morgen_random,
        "Bis.": end_morgen_random,
        "Von": von_nacht_random,
        "Bis": end_nacht_random,
        "Ges.Stunden": gesamt_stunden_random})
    df.columns = pd.MultiIndex.from_tuples(zip([f"Mitarber: {name}", "", "", "", "", "",""], df.columns))
    
    df_all.append(df)
    
    

def dfs_tabs(df_list, sheet_list):

    output = io.BytesIO()

    writer = pd.ExcelWriter(output, engine='xlsxwriter')   
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0 , startcol=0)   
    writer.close()

    processed_data = output.getvalue()
    return processed_data

excel_files = dfs_tabs(df_all, all_names)

        
if st.button("下载Excel"):        
        st.download_button(
            label="Download Excel",
            data=excel_files,
            file_name="arbeitstunden.xlsx",
            mime="application/vnd.ms-excel",
        )


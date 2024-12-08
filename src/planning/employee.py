import os
from datetime import timedelta

import numpy as np
import pandas as pd
import yaml
from openpyxl import load_workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)


def create_employees(employee_restrictions):
    """Create employees.

    :param employee_restrictions:
    :return:
    """
    # Employees information
    employees_full = [
        {
            "capacity": 1,
        }
    ]

    employees_partial = [
        {
            "capacity": 0.77,
        }
    ]

    employees_temp = employees_full * 2 + employees_partial * 1

    employees = []
    for index, one_employee in enumerate(employees_temp):
        employees.append(
            {
                "name": f"E{index + 1}",
                "capacity": one_employee["capacity"],
                "max_hours_year": employee_restrictions["max_hours_year_employee"]
                * one_employee["capacity"],
                "max_hours_week": employee_restrictions["max_hours_week_employee"]
                * one_employee["capacity"],
            }
        )

    return employees


def create_employees_with_dates(start_date, num_days, employees):
    """Create employees with dates.

    :param start_date:
    :param num_days:
    :param employees:
    :return:
    """
    dates = pd.date_range(start=start_date, periods=num_days, freq="D")
    employees_info = pd.DataFrame(
        index=dates, columns=[emp["name"] for emp in employees], data=""
    )
    return employees_info, dates


def init_employees_by_shifts(dates, employee_restrictions):
    """Init employees by shifts.

    :param dates:
    :param employee_restrictions:
    :return:
    """
    all_employees_by_shift = pd.DataFrame(
        index=dates,
        columns=[one_shift for one_shift in employee_restrictions["shifts"]],
    )
    all_employees_by_shift[:] = 0

    return all_employees_by_shift


def load_data_by_date(
    all_employees_by_shift, employee_restrictions, employees_info, employees
):
    """Load data by date.

    :param all_employees_by_shift:
    :param employee_restrictions:
    :param employees_info:
    :param employees:
    :return:
    """
    for date in all_employees_by_shift.index:
        for shift in all_employees_by_shift.columns:
            if (
                all_employees_by_shift.loc[date, shift]
                >= employee_restrictions["max_persons_per_shift"][shift]
            ):  # No more employees needed
                continue

            available_employees = []

            for employee in employees_info.columns:
                if (
                    not employees_info.loc[date, employee]
                    or employees_info.loc[date, employee] == "-"
                ):
                    six_days_ago = date - timedelta(days=6)
                    five_days_ago = date - timedelta(days=5)
                    thirty_days_ago = date - timedelta(days=30)
                    yesterday = date - timedelta(days=1)

                    last_6_days_employee = employees_info.loc[
                        six_days_ago:date, employee
                    ]
                    last_5_days_employee = employees_info.loc[
                        five_days_ago:date, employee
                    ]
                    last_30_days_employee = employees_info.loc[
                        thirty_days_ago:date, employee
                    ]
                    if yesterday in employees_info.index:
                        value_yesterday = employees_info.loc[yesterday, employee]
                    else:
                        value_yesterday = None

                    total_worked_days_in_6_days = (
                        last_6_days_employee.value_counts()
                        .reindex(["M", "T"], fill_value=0)
                        .sum()
                    )
                    total_rest_days_in_5_days = (
                        last_5_days_employee.value_counts()
                        .reindex(["-"], fill_value=0)
                        .sum()
                    )

                    weekend_days = last_30_days_employee[
                        last_30_days_employee.index.weekday.isin([6])
                    ]
                    total_worked_weekends_in_30_days = (
                        weekend_days.value_counts()
                        .reindex(["M", "T"], fill_value=0)
                        .sum()
                    )

                    sum_m_t_per_column = employees_info.apply(
                        lambda col: col.isin(["M", "T"]).sum(), axis=0
                    )
                    total_sum_m_t = employees_info[employee].isin(["M", "T"]).sum()

                    employee_capacity = next(
                        emp["capacity"] for emp in employees if emp["name"] == employee
                    )
                    if (
                        (
                            (total_sum_m_t * employee_restrictions["hours_per_shift"])
                            >= (
                                employee_restrictions["max_hours_year_employee"]
                                * employee_capacity
                            )
                        )
                        or (
                            (
                                (total_worked_days_in_6_days + 1)
                                * employee_restrictions["hours_per_shift"]
                            )
                            >= (
                                employee_restrictions["max_hours_week_employee"]
                                * employee_capacity
                            )
                        )
                        or (
                            value_yesterday
                            and (value_yesterday == "T" and shift == "M")
                        )
                    ):
                        employees_info.loc[date, employee] = ""
                    else:
                        available_employees.append(
                            {
                                "employee": employee,
                                "total_worked_weekends_in_30_days": total_worked_weekends_in_30_days,
                            }
                        )

            # Sort available_employees per total_worked_weekends_in_30_days - Descending order
            available_employees = sorted(
                available_employees,
                key=lambda x: x["total_worked_weekends_in_30_days"],
                reverse=False,
            )
            for one_employee in available_employees:
                all_employees_by_shift.loc[date, shift] += 1
                employees_info.loc[date, one_employee["employee"]] = shift
                if (
                    all_employees_by_shift.loc[date, shift]
                    >= employee_restrictions["max_persons_per_shift"][shift]
                ):  # No more employees needed
                    break
        for one_employee in employees_info.columns:
            if employees_info.loc[date, one_employee] == "":
                employees_info.loc[date, one_employee] = "-"


def modify_index_to_datetime(dataframe):
    """Modify dataframe index to datetime.

    :param dataframe:
    :return:
    """
    dataframe.index = pd.to_datetime(dataframe.index)
    dataframe.index = dataframe.index.strftime("%Y-%m-%d")


def generate_excel(dataframe, filename):
    """Generate excel.

    :param dataframe:
    :param filename:
    :return:
    """
    output_filename = filename
    dataframe.to_excel(output_filename, sheet_name="Shift Schedule")


def load_translations():
    """Load translations.

    :return:
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lang_file_path = os.path.join(script_dir, "lang.yaml")
    with open(lang_file_path) as file:
        lang_data = yaml.safe_load(file)
    return lang_data


def create_transposed_dataframe(employees_info, lang="es"):
    """Create a transposed dataframe.

    :param employees_info:
    :param lang:
    :return:
    """
    employees_info.index = pd.to_datetime(employees_info.index)

    lang_data = load_translations()

    day_of_month = employees_info.index.day
    days_of_week_map = {i: lang_data["days_of_week"][i][lang] for i in range(7)}
    months_map = {month: lang_data["months"][month][lang] for month in range(1, 13)}
    month = employees_info.index.month.map(months_map)
    day_of_week = employees_info.index.dayofweek.map(days_of_week_map)

    multi_index_index = pd.MultiIndex.from_arrays(
        [month, day_of_week, day_of_month], names=["", "", ""]
    )

    employees_info.index = multi_index_index

    transposed_employees_info = employees_info.T

    return transposed_employees_info


def generate_summary(employees, employee_restrictions, transposed_employees_info):
    """Generate summary.

    :param employees:
    :param employee_restrictions:
    :param transposed_employees_info:
    :return:
    """
    transposed_employees_info["THT"] = transposed_employees_info.apply(
        lambda row: (row.value_counts().get("M", 0) + row.value_counts().get("T", 0))
        * employee_restrictions["hours_per_shift"],
        axis=1,
    )

    transposed_employees_info["MH"] = transposed_employees_info.index.map(
        lambda emp: next(
            employee["max_hours_year"]
            for employee in employees
            if employee["name"] == emp
        )
    )
    transposed_employees_info["Diff"] = (
        transposed_employees_info["MH"] - transposed_employees_info["THT"]
    )

    sum_m_t = transposed_employees_info.apply(
        lambda col: col.isin(["M", "T"]).sum(), axis=0
    )
    new_row = pd.Series(sum_m_t, name="Total")
    transposed_employees_info = pd.concat(
        [transposed_employees_info, new_row.to_frame().T]
    )

    transposed_employees_info.loc["Total", ["THT", "MH", "Diff"]] = [np.nan] * 3

    return transposed_employees_info


def generate_transposed_excel_with_styles(
    transposed_employees_info, employee_restrictions, filename
):
    """Generate transposed excel with styles.

    :param transposed_employees_info:
    :param employee_restrictions:
    :param filename:
    :return:
    """
    output_filename = filename
    transposed_employees_info.to_excel(output_filename, sheet_name="Shift Schedule")

    workbook = load_workbook(output_filename)
    worksheet = workbook["Shift Schedule"]

    worksheet.delete_rows(4)

    min_width = 3
    for col in worksheet.iter_cols():
        for cell in col:
            if not any(
                cell.coordinate in merged_cell
                for merged_cell in worksheet.merged_cells.ranges
            ):
                column = cell.column_letter
                worksheet.column_dimensions[column].width = min_width
                break

    for i in range(0, 3):
        column_letter = worksheet.cell(
            row=1, column=worksheet.max_column - i
        ).column_letter
        worksheet.column_dimensions[column_letter].width = 7
        for cell in worksheet[column_letter]:
            cell.alignment = Alignment(horizontal="center")

    first_column_letter = worksheet.cell(row=1, column=1).column_letter
    worksheet.column_dimensions[first_column_letter].width = 7

    fill = PatternFill(start_color="0099FF", end_color="0099FF", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)

    for cell in worksheet[1]:
        cell.fill = fill
        cell.font = font

    weekend_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col in worksheet.iter_cols(
        min_row=2, max_row=worksheet.max_row, min_col=2, max_col=worksheet.max_column
    ):
        day_of_week_cell = col[0]
        if day_of_week_cell.value in ["S", "D"]:
            for cell in col:
                cell.fill = weekend_fill

    for row in worksheet.iter_rows(
        min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column
    ):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

    min_persons_day = (
        employee_restrictions["min_persons_per_shift"]["M"]
        + employee_restrictions["min_persons_per_shift"]["T"]
    )

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    total_row = worksheet.max_row
    for cell in worksheet[total_row]:
        if str(cell.value).isdigit() and int(cell.value) < min_persons_day:
            cell.fill = red_fill

    workbook.save(output_filename)

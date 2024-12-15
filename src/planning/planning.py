# Planning script. This script is used to generate the planning for the employees.

import os

from config import (
    employee_restrictions,
)
from employee import (
    assign_vacations,
    create_employees,
    create_employees_with_dates,
    create_transposed_dataframe,
    generate_excel,
    generate_summary,
    generate_transposed_excel_with_styles,
    init_employees_by_shifts,
    load_data_by_date,
    modify_index_to_datetime,
)


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    samples_dir = os.path.join(script_dir, "samples")

    year = 2025
    start_date = f"{year}-01-01"
    employees = create_employees(employee_restrictions)
    employees_info, dates = create_employees_with_dates(start_date, 365, employees)
    all_employees_by_shift = init_employees_by_shifts(dates, employee_restrictions)
    assign_vacations(employees_info)
    load_data_by_date(all_employees_by_shift, employee_restrictions, employees_info, employees, start_date)
    modify_index_to_datetime(all_employees_by_shift)
    modify_index_to_datetime(employees_info)

    generate_excel(employees_info, f"samples/{year}_employees_0.xlsx")

    transposed_employees_info = create_transposed_dataframe(employees_info)
    transposed_employees_info = generate_summary(employees, employee_restrictions, transposed_employees_info)

    generate_transposed_excel_with_styles(
        transposed_employees_info, employee_restrictions, f"{samples_dir}/{year}_planning_0.xlsx"
    )


if __name__ == "__main__":
    main()

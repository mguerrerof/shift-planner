"""This module handles the planning and scheduling of employee vacations and shifts.

Functions:
    main(): The main function that orchestrates the creation of employee schedules,
            assignment of vacations, and generation of Excel reports.

Modules imported:
    os: Provides a way of using operating system dependent functionality.
    config: Contains employee restrictions configuration.
    employee: Contains functions for handling employee data and generating reports.

The main function performs the following steps:
    1. Sets up the script and output directories.
    2. Initializes the year and start date for the planning.
    3. Creates employee objects based on restrictions.
    4. Generates employee information and dates.
    5. Initializes employees by shifts.
    6. Assigns vacations to employees.
    7. Loads data by date for all employees by shift.
    8. Modifies the index of dataframes to datetime.
    9. Generates an Excel file with employee information.
    10. Creates a transposed dataframe of employee information.
    11. Generates a summary of the transposed employee information.
    12. Generates a styled Excel file with the transposed and summarized employee information.
"""

import os

from employee import (
    assign_vacations,
    create_employees_with_dates,
    create_transposed_dataframe,
    generate_excel,
    generate_summary,
    generate_transposed_excel_with_styles,
    init_employees_by_shifts,
    load_config,
    load_data_by_date,
    load_employees_from_yaml,
    modify_index_to_datetime,
)


def main():
    year = 2025
    case = "case_1"
    script_dir = os.path.abspath("../../")
    output_dir = os.path.join(script_dir, "output")
    output_file = os.path.join(output_dir, str(year), case, "generated_from_script.xlsx")
    employees_file = os.path.join(script_dir, "data", "2025", case, "employees.yaml")
    vacations_file = os.path.join(script_dir, "data", "2025", case, "vacations.yaml")
    config_file = os.path.join(script_dir, "data", "2025", case, "config.json")

    config = load_config(config_file)
    employee_restrictions = config["employee_restrictions"]

    start_date = f"{year}-01-01"
    employees = load_employees_from_yaml(employees_file, employee_restrictions)
    employees_info, dates = create_employees_with_dates(start_date, 365, employees)
    all_employees_by_shift = init_employees_by_shifts(dates, employee_restrictions)
    assign_vacations(employees_info, vacations_file)
    load_data_by_date(all_employees_by_shift, employee_restrictions, employees_info, employees, start_date)
    modify_index_to_datetime(all_employees_by_shift)
    modify_index_to_datetime(employees_info)

    generate_excel(employees_info, output_file)

    transposed_employees_info = create_transposed_dataframe(employees_info)
    transposed_employees_info = generate_summary(employees, employee_restrictions, transposed_employees_info)

    generate_transposed_excel_with_styles(transposed_employees_info, employee_restrictions, output_file)


if __name__ == "__main__":
    main()

# planning
Planning Management

## Install 

**Python version:** 3.13.0

Execute the following command:
```
pip install -r requirements.txt
````

### Num employees in each shift

![alt text](doc/all_employees_in_shift.png)

### 3 employees:

![alt text](doc/3_employees_with_2_shifts.png)

### 30 days with 3 employees:

![alt text](doc/30_days.png)

### Sample with a transpose table
![alt text](doc/transpose_table.png)

### Export to Excel
```
output_filename = "../samples/m_a_2025.xlsx"
employees_info.index = pd.to_datetime(employees_info.index)
employees_info.index = employees_info.index.strftime("%Y-%m-%d")
employees_info.to_excel(output_filename, sheet_name="Shift Schedule")
```
import pandas as pd

def analyze_excel_file(file_path):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)

    # Convert the 'Pay Cycle Start Date' and 'Pay Cycle End Date' columns to datetime format
    df['Pay Cycle Start Date'] = pd.to_datetime(df['Pay Cycle Start Date'])
    df['Pay Cycle End Date'] = pd.to_datetime(df['Pay Cycle End Date'])

    # Function to check consecutive days worked
    def consecutive_days(series):
        return all((d2 - d1).days == 1 for d1, d2 in zip(series, series[1:]))

    # Function to check time between shifts
    def time_between_shifts(series):
        return all(1 < (t2 - t1).seconds // 3600 < 10 for t1, t2 in zip(series, series[1:]))

    # Function to check single shift duration
    def long_shift(series):
        return any((end_time - start_time).seconds // 3600 > 14 for start_time, end_time in zip(series, series[1:]))

    # Group by 'Employee Name' and apply the functions
    SevenConsecutiveDays = df.groupby('Employee Name')['Pay Cycle Start Date'].agg(consecutive_days)
    OneToTenHoursShift = df.groupby('Employee Name')['Time'].agg(time_between_shifts)
    MoreThan14Shift = df.groupby('Employee Name')['Time'].agg(long_shift)

    # Get 'Position ID' for each employee
    position_id_mapping = df.groupby('Employee Name')['Position ID'].first()

    # Print results
    print("a. Employees who have worked for 7 consecutive days:")
    print_result(SevenConsecutiveDays[SevenConsecutiveDays].index, position_id_mapping)

    print("\nb. Employees who have less than 10 hours between shifts but greater than 1 hour:")
    print_result(OneToTenHoursShift[OneToTenHoursShift].index, position_id_mapping)

    print("\nc. Employees who have worked for more than 14 hours in a single shift:")
    print_result(MoreThan14Shift[MoreThan14Shift].index, position_id_mapping)

def print_result(employees, position_id_mapping):
    # Print the results with 'Position ID'
    result_df = pd.DataFrame({'Position': position_id_mapping[employees], 'Employee Name': employees})
    print(result_df.to_string(index=False))

analyze_excel_file('C:\\Users\\RITIKA\\Desktop\\DHRUV\\Projects\\Assignment\\Assignment_Timecard.xlsx')
import openpyxl
from datetime import datetime, timedelta


def to_rack(filtered_dates, current_date, data_dict, filename="to_rack.txt"):
    """
    Filters and writes entries to a text file for arrivals 3+ days away.

    Args:
        filtered_dates (list): List of filtered arrival dates.
        current_date (datetime): Current date and time.
        data_dict (dict): Dictionary containing parking information.
        filename (str, optional): Name of the output text file. Defaults to "to_rack.txt".
    """

    filtered_entries = [
        (name, loc, num, date)
        for name, (loc, num, date) in data_dict.items()
        if date >= current_date + timedelta(days=3) and loc != "RACK" and "T3" not in loc
    ]

    # Open the file in write mode ("w") and close it automatically
    with open(filename, "w") as f:
        for name, loc, num, date in filtered_entries:
            f.write(
                f"Last Name: {name}, Parked Location: {loc}, Reservation Number: {num}, Arrival Date: {date.strftime('%Y-%m-%d %H:%M:%S')}\n"
            )


def from_rack(filtered_dates, current_date, data_dict, filename="from_rack.txt"):
    """
    Filters and writes entries to a text file for arrivals today or tomorrow.

    Args:
        filtered_dates (list): List of filtered arrival dates.
        current_date (datetime): Current date and time.
        data_dict (dict): Dictionary containing parking information.
        filename (str, optional): Name of the output text file. Defaults to "from_rack.txt".
    """

    filtered_entries = [
        (name, loc, num, date)
        for name, (loc, num, date) in data_dict.items()
        if current_date <= date < current_date + timedelta(days=1.5) and "T3" not in loc and "T4" not in loc
    ]

    # Open the file in write mode ("w") and close it automatically
    with open(filename, "w") as f:
        for name, loc, num, date in filtered_entries:
            f.write(
                f"Last Name: {name}, Parked Location: {loc}, Reservation Number: {num}, Arrival Date: {date.strftime('%Y-%m-%d %H:%M:%S')}\n"
            )
            
def write_to_file(filtered_dates, current_date, data_dict, filename="parking_report.txt"):
    """
    Filters and writes entries to a text file with specified format.

    Args:
        filtered_dates (list): List of filtered arrival dates.
        current_date (datetime): Current date and time.
        data_dict (dict): Dictionary containing parking information.
        filename (str, optional): Name of the output text file. Defaults to "parking_report.txt".
    """

    with open(filename, "a") as f:  # Open in append mode ("a")
        # To Rack entries
        f.write("TO RACK:\n")
        for name, (loc, num, date) in data_dict.items():
            if date >= current_date + timedelta(days=3) and loc != "RACK" and "T3" not in loc:
                f.write(f"{date.strftime('%Y-%m-%d %H:%M:%S')}, {name}, {num}, {loc}\n")

        # From Rack entries
        f.write("\nFROM RACK:\n")
        for name, (loc, num, date) in data_dict.items():
            if current_date <= date < current_date + timedelta(days=1.5) and "T3" not in loc and "T4" not in loc:
                f.write(f"{date.strftime('%Y-%m-%d %H:%M:%S')}, {name}, {num}, {loc}\n")
            
def main():
    wb = openpyxl.load_workbook('ParkedLocationInventory.xlsx')
    sheet = wb['ParkedLocationInventory']

    arrival_dates = sheet['L']
    parked_location = sheet['A']
    reservation_num = sheet['E']
    last_name = sheet['G']

    # Removes None and Header values
    filtered_dates = [cell.value for cell in arrival_dates if cell.value not in (None, 'Arrival Scheduled')]
    filtered_parked_location = [cell.value for cell in parked_location if cell.value not in (None, 'Parked Location')]
    filtered_reservation_num = [cell.value for cell in reservation_num if cell.value not in (None, 'Reservation#')]
    filtered_last_name = [cell.value for cell in last_name if cell.value not in (None, 'Last Name')]

    data_dict = {}

    for i in range(len(filtered_last_name)):
        last_name = filtered_last_name[i]
        parked_loc = filtered_parked_location[i]
        res_num = filtered_reservation_num[i]
        arrival_date = filtered_dates[i]

        data_dict[last_name] = (parked_loc, res_num, arrival_date)

    # Call the function to process the data
    #to_rack(filtered_dates, datetime.now(), data_dict)
    #from_rack(filtered_dates, datetime.now(), data_dict)
    
    write_to_file(filtered_dates, datetime.now(), data_dict)


# Call the main function to execute the program
if __name__ == "__main__":
    main()

from zipfile import ZipFile
import pandas as pd
import os

# read your Excel file
excel_file_name = "Book1.xlsx"

# add option for headers if there are any
excel_data = pd.read_excel(excel_file_name, header=[0])
# excel_data = pd.read_excel(excel_file_name, header=None)

# loop through the rows
# mine was odd - IP address, even - switch name(location)
for i in range(0, len(excel_data), 2):
    switch_location = excel_data.iloc[i+1,0]
    ip_address = excel_data.iloc[i,0]
    # prepare your .bat file naming convention
    file_name = f"{switch_location} {ip_address }.bat"
    # prepare the ping command
    message = f"ping {ip_address} -t"

    # write the command into the newly generated file
    with open(file_name, "w") as file:
        file.write(message)

    # optional output message for successful file creation
    print(f"{file_name} created")

# zip all the created files into one for easy distribution
with ZipFile("all_switches.zip", "w") as z:
    for file_name in os.listdir():
        if file_name.endswith(".bat"):
            z.write(file_name)
